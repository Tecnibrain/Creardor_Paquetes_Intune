#Requires -Version 5.1
<#
.SYNOPSIS
    Fábrica de Paquetes PSADT para Intune (versión 4.0).

.DESCRIPTION
    Herramienta con interfaz gráfica para crear paquetes PSAppDeployToolkit 4.x listos para Intune
    a partir de un instalador MSI, EXE o un script PS1, sin escribir código.

    Todo es automático:
      - Primera ejecución: consigue el módulo PSADT (carpeta Herramientas, módulo instalado o
        descarga de la PowerShell Gallery si hay Internet), crea la plantilla corporativa en
        Recursos y genera logo y banner.
      - Pestaña "Plantilla corporativa": empresa, color, logo, idioma, estilo de ventanas,
        ruta de logs, aplazamientos y mensaje final por defecto. Se aplica a todos los paquetes.
      - Pestaña "Crear paquete": eliges el instalador, los datos se llenan solos y la
        herramienta genera el paquete, el script de despliegue, la detección, la ficha de
        Intune, el logo de la app y el .intunewin.

    Compatible con las plantillas PSADT 4.0 / 4.1 (funciones Install-ADTDeployment) y
    4.2 (bloques New-Variable -Name Install).

    Modo web (-ModoWeb): la misma Fábrica en el navegador. Levanta un servidor mínimo SOLO en
    127.0.0.1 (no acepta conexiones de otros equipos), con un token aleatorio por sesión, y abre
    Edge. No necesita permisos de administrador, IIS ni Internet. Usa las mismas funciones que la
    interfaz de escritorio.

.EXAMPLE
    .\Nuevo-Paquete.ps1
    Abre la interfaz gráfica.

.EXAMPLE
    .\Nuevo-Paquete.ps1 -ModoWeb
    Abre la Fábrica en el navegador (http://127.0.0.1:<puerto libre>).

.EXAMPLE
    .\Nuevo-Paquete.ps1 -SinInterfaz -Instalador 'D:\Fuentes\App.msi' -ProcesosACerrar 'app' -GenerarIntunewin

.NOTES
    Windows PowerShell 5.1. No requiere permisos de administrador para crear paquetes.
#>
[CmdletBinding()]
param(
    [string]$Instalador,
    [string]$Fabricante,
    [string]$NombreApp,
    [string]$Version,
    [string]$Descripcion,
    [ValidateSet('x64', 'x86')]
    [string]$Arquitectura = 'x64',
    [string]$ArgumentosInstalacion,
    [string]$DesinstaladorExe,
    [string]$ArgumentosDesinstalacion,
    [string]$NombreDeteccion,
    [string]$ProcesosACerrar,
    [switch]$SinAplazamiento,
    [ValidateRange(1, 99)]
    [int]$VecesAplazar = 3,
    [string]$MensajeFinal,
    [string]$MensajeBienvenida,
    [string]$TextoBotonInstalar,
    [string]$TextoBotonAplazar,
    [ValidateRange(0, 365)]
    [int]$DiasLimite = 0,
    [ValidateRange(0, 240)]
    [int]$CuentaRegresivaMin = 0,
    [switch]$SiemprePreguntar,
    [string]$LogoApp,
    [switch]$LogoAppEnVentanas,
    [switch]$CopiarCarpetaCompleta,
    [switch]$GenerarIntunewin,
    [string]$CarpetaDestino,
    [switch]$IncluirFuente,
    [switch]$SinInterfaz,
    [switch]$ModoWeb,
    [ValidateRange(0, 65535)]
    [int]$Puerto = 0,
    [switch]$SinNavegador
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

#region ===================== Variables globales =====================
$script:Raiz            = $PSScriptRoot
$script:Herramientas    = Join-Path $script:Raiz 'Herramientas'
$script:Recursos        = Join-Path $script:Raiz 'Recursos'
$script:CarpetaPaquetes = Join-Path $script:Raiz 'Paquetes'
$script:CarpetaSalida   = Join-Path $script:Raiz 'Salida'
$script:RutaConfigJson  = Join-Path $script:Recursos 'Corporativo.json'
$script:CajaRegistro    = $null
$script:TextoEstado     = $null
$script:Utf8ConBom      = New-Object System.Text.UTF8Encoding($true)
$script:Utf8SinBom      = New-Object System.Text.UTF8Encoding($false)
$script:RutaHistorial   = Join-Path $script:Recursos 'Historial.json'
$script:RutaEsteScript  = $PSCommandPath
$script:VersionFabrica  = '4.0'
$script:RegistroWeb     = $null   # en modo web: lista de líneas del registro que se envían al navegador
$script:SecuenciaRegistro = 0
$script:LargoPlantilla  = 0

$script:ConfigPorDefecto = [ordered]@{
    Empresa       = 'Bancolombia'
    ColorAcento   = '#2C2A29'
    EstiloDialogo = 'Fluent'
    Idioma        = 'es'
    RutaLogs      = '$envWinDir\Logs\Software'
    VecesAplazar  = 3
    MensajeFinal  = ''
    Logo          = ''
    CuentaRegresivaMin = 0
    DiasLimite         = 0
    CarpetaDestino     = ''
    Banner             = ''
    HuellasAssets      = ''
    MensajeBienvenida  = ''
    TextoBotonInstalar = ''
    TextoBotonAplazar  = ''
}

$script:VersionPSADT = [version]'0.0'

$script:OrdenFases = @('Pre-Install', 'Install', 'Post-Install', 'Pre-Uninstall', 'Uninstall', 'Post-Uninstall', 'Pre-Repair', 'Repair', 'Post-Repair')
#endregion

#region ===================== Utilidades =====================
function Write-Registro {
    param(
        [Parameter(Mandatory)][string]$Mensaje,
        [ValidateSet('Info', 'OK', 'Aviso', 'Error')][string]$Tipo = 'Info'
    )
    $prefijo = switch ($Tipo) { 'OK' { '[OK]   ' } 'Aviso' { '[AVISO]' } 'Error' { '[ERROR]' } default { '[INFO] ' } }
    $linea = '{0} {1} {2}' -f (Get-Date -Format 'HH:mm:ss'), $prefijo, $Mensaje
    if ($null -ne $script:RegistroWeb) {
        # Modo web: se guarda para enviarlo al navegador en la respuesta de la petición en curso.
        $script:SecuenciaRegistro++
        $script:RegistroWeb.Add(@{ Seq = $script:SecuenciaRegistro; Linea = $linea; Tipo = $Tipo; Mensaje = $Mensaje })
        if ($script:RegistroWeb.Count -gt 2000) { $script:RegistroWeb.RemoveRange(0, 500) }
    }
    if ($null -ne $script:CajaRegistro) {
        # Interfaz WPF: agrega la línea, muestra el estado y deja que la ventana se redibuje.
        $script:CajaRegistro.AppendText($linea + [Environment]::NewLine)
        $script:CajaRegistro.ScrollToEnd()
        if ($null -ne $script:TextoEstado) {
            $script:TextoEstado.Text = $Mensaje
            $script:TextoEstado.Foreground = switch ($Tipo) {
                'OK' { [System.Windows.Media.Brushes]::SeaGreen }
                'Aviso' { [System.Windows.Media.Brushes]::DarkOrange }
                'Error' { [System.Windows.Media.Brushes]::Firebrick }
                default { [System.Windows.Media.Brushes]::DimGray }
            }
        }
        $script:CajaRegistro.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
    }
    else {
        $color = switch ($Tipo) { 'OK' { 'Green' } 'Aviso' { 'Yellow' } 'Error' { 'Red' } default { 'Gray' } }
        Write-Host $linea -ForegroundColor $color
    }
}

function ConvertTo-TextoPS {
    # Cadena PowerShell entre comillas simples (escapa ').
    param([AllowEmptyString()][string]$Texto)
    return "'" + ($Texto -replace "'", "''") + "'"
}

function ConvertTo-TextoPSExpandible {
    # Cadena entre comillas dobles: permite variables como $env:ProgramFiles.
    param([AllowEmptyString()][string]$Texto)
    $escapado = $Texto -replace '`', '``' -replace '"', '`"'
    return '"' + $escapado + '"'
}

function Get-NombreLimpio {
    param([string]$Texto)
    $t = $Texto.Normalize([System.Text.NormalizationForm]::FormD)
    $t = $t -replace '\p{Mn}', ''
    $t = $t -replace '\s+', ''
    $t = $t -replace '[^A-Za-z0-9\.\-_]', ''
    return $t
}

function Get-Iniciales {
    param([string]$Texto)
    $palabras = @(($Texto -split '[\s_\-\.]+') | Where-Object { $_ -match '^[A-Za-z0-9]' })
    if ($palabras.Count -ge 2) { return ("{0}{1}" -f $palabras[0][0], $palabras[1][0]).ToUpper() }
    if ($palabras.Count -eq 1) { return $palabras[0].Substring(0, [Math]::Min(2, $palabras[0].Length)).ToUpper() }
    return 'AP'
}

function ConvertTo-Indentado {
    param([string[]]$Lineas, [int]$Espacios = 4)
    $prefijo = ' ' * $Espacios
    $salida = @()
    foreach ($l in $Lineas) { if ($l -eq '') { $salida += '' } else { $salida += ($prefijo + $l) } }
    return , $salida
}
#endregion

#region ===================== Configuración corporativa =====================
function Get-ConfigCorporativa {
    $cfg = [ordered]@{}
    foreach ($k in $script:ConfigPorDefecto.Keys) { $cfg[$k] = $script:ConfigPorDefecto[$k] }
    if (Test-Path -LiteralPath $script:RutaConfigJson) {
        try {
            $obj = [System.IO.File]::ReadAllText($script:RutaConfigJson) | ConvertFrom-Json
            foreach ($p in $obj.PSObject.Properties) {
                if ($cfg.Contains($p.Name) -and $null -ne $p.Value) { $cfg[$p.Name] = $p.Value }
            }
        }
        catch {
            Write-Registro -Mensaje "No se pudo leer Corporativo.json; se usan valores por defecto. $($_.Exception.Message)" -Tipo Aviso
        }
    }
    return $cfg
}

function Save-ConfigCorporativa {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Cfg)
    $json = $Cfg | ConvertTo-Json
    [System.IO.File]::WriteAllText($script:RutaConfigJson, $json, $script:Utf8ConBom)
}

function Set-ValorConfig {
    # Cambia una clave de config.psd1. -SoloEntreComillas: solo reemplaza si el valor actual va entre comillas simples.
    param([string]$Texto, [string]$Clave, [string]$ValorLiteral, [switch]$SoloEntreComillas)
    if ($SoloEntreComillas) { $patron = "(?m)^(\s*$Clave\s*=\s*)'[^']*'" }
    else { $patron = "(?m)^(\s*$Clave\s*=\s*)(\`$null|'[^']*'|\S+)" }
    $rx = New-Object System.Text.RegularExpressions.Regex($patron)
    if (-not $rx.IsMatch($Texto)) {
        Write-Registro -Mensaje "config.psd1: no se encontró '$Clave'; se deja como está." -Tipo Aviso
        return $Texto
    }
    $reemplazo = '${1}' + $ValorLiteral.Replace('$', '$$')
    return $rx.Replace($Texto, $reemplazo, 1)
}

function ConvertFrom-ColorHex {
    param([string]$Hex)
    $h = $Hex.Trim().TrimStart('#')
    if ($h -notmatch '^[0-9A-Fa-f]{6}$') { $h = '2C2A29' }
    return [System.Drawing.Color]::FromArgb(255, [Convert]::ToInt32($h.Substring(0, 2), 16), [Convert]::ToInt32($h.Substring(2, 2), 16), [Convert]::ToInt32($h.Substring(4, 2), 16))
}

function Get-ColorTexto {
    param([System.Drawing.Color]$Color)
    $luminancia = (0.299 * $Color.R + 0.587 * $Color.G + 0.114 * $Color.B) / 255
    if ($luminancia -gt 0.6) { return [System.Drawing.Color]::FromArgb(255, 30, 30, 30) }
    return [System.Drawing.Color]::White
}
#endregion

#region ===================== Imágenes (logo, banner, icono de la app) =====================
function Get-ImagenSinBloqueo {
    # Carga una imagen sin dejar el archivo bloqueado.
    param([string]$Ruta)
    if ([string]::IsNullOrWhiteSpace($Ruta) -or -not (Test-Path -LiteralPath $Ruta)) { return $null }
    $bytes = [System.IO.File]::ReadAllBytes($Ruta)
    $ms = New-Object System.IO.MemoryStream(, $bytes)
    try {
        $tmp = [System.Drawing.Image]::FromStream($ms)
        try { $copia = New-Object System.Drawing.Bitmap($tmp) } finally { $tmp.Dispose() }
    }
    finally { $ms.Dispose() }
    return $copia
}

function New-RectanguloRedondeado {
    param([single]$X, [single]$Y, [single]$Ancho, [single]$Alto, [single]$Radio)
    $ruta = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = [single]($Radio * 2)
    $ruta.AddArc($X, $Y, $d, $d, 180, 90)
    $ruta.AddArc($X + $Ancho - $d, $Y, $d, $d, 270, 90)
    $ruta.AddArc($X + $Ancho - $d, $Y + $Alto - $d, $d, $d, 0, 90)
    $ruta.AddArc($X, $Y + $Alto - $d, $d, $d, 90, 90)
    $ruta.CloseFigure()
    return $ruta
}

function Invoke-DibujoAjustado {
    # Dibuja una imagen dentro de un área conservando proporción y centrada.
    param($Grafico, [System.Drawing.Image]$Imagen, [single]$X, [single]$Y, [single]$Ancho, [single]$Alto)
    $escala = [Math]::Min($Ancho / $Imagen.Width, $Alto / $Imagen.Height)
    $w = [single]($Imagen.Width * $escala)
    $h = [single]($Imagen.Height * $escala)
    $Grafico.DrawImage($Imagen, [single]($X + ($Ancho - $w) / 2), [single]($Y + ($Alto - $h) / 2), $w, $h)
}

function Initialize-Grafico {
    param($Grafico)
    $Grafico.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $Grafico.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $Grafico.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $Grafico.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
}

function New-IconoPng {
    # Icono cuadrado: logo de la empresa (si hay) o iniciales sobre el color corporativo.
    param(
        [Parameter(Mandatory)][string]$Ruta,
        [int]$Ancho = 256,
        [int]$Alto = 256,
        [string]$ColorHex = '#2C2A29',
        [string]$Iniciales = 'AP',
        [System.Drawing.Image]$ImagenOrigen
    )
    $bmp = New-Object System.Drawing.Bitmap($Ancho, $Alto, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        Initialize-Grafico $g
        $g.Clear([System.Drawing.Color]::Transparent)
        $lado = [single][Math]::Min($Ancho, $Alto)
        $x = [single](($Ancho - $lado) / 2)
        $y = [single](($Alto - $lado) / 2)
        if ($null -ne $ImagenOrigen) {
            Invoke-DibujoAjustado -Grafico $g -Imagen $ImagenOrigen -X $x -Y $y -Ancho $lado -Alto $lado
        }
        else {
            $color = ConvertFrom-ColorHex $ColorHex
            $forma = New-RectanguloRedondeado -X $x -Y $y -Ancho $lado -Alto $lado -Radio ($lado * 0.22)
            $brocha = New-Object System.Drawing.SolidBrush($color)
            $brochaTexto = New-Object System.Drawing.SolidBrush((Get-ColorTexto $color))
            $fuente = New-Object System.Drawing.Font('Segoe UI', [single]($lado * 0.40), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
            $formato = New-Object System.Drawing.StringFormat
            $formato.Alignment = [System.Drawing.StringAlignment]::Center
            $formato.LineAlignment = [System.Drawing.StringAlignment]::Center
            try {
                $g.FillPath($brocha, $forma)
                $g.DrawString($Iniciales, $fuente, $brochaTexto, (New-Object System.Drawing.RectangleF($x, $y, $lado, $lado)), $formato)
            }
            finally {
                $formato.Dispose(); $fuente.Dispose(); $brochaTexto.Dispose(); $brocha.Dispose(); $forma.Dispose()
            }
        }
        $bmp.Save($Ruta, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $g.Dispose()
        $bmp.Dispose()
    }
}

function New-BannerPng {
    # Banner horizontal (estilo Classic): color corporativo, logo o iniciales a la izquierda y nombre de la empresa.
    param(
        [Parameter(Mandatory)][string]$Ruta,
        [int]$Ancho = 900,
        [int]$Alto = 125,
        [string]$ColorHex = '#2C2A29',
        [string]$Empresa = '',
        [string]$Iniciales = 'AP',
        [System.Drawing.Image]$ImagenOrigen
    )
    $bmp = New-Object System.Drawing.Bitmap($Ancho, $Alto, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $color = ConvertFrom-ColorHex $ColorHex
    $colorTexto = Get-ColorTexto $color
    $brochaFondo = New-Object System.Drawing.SolidBrush($color)
    $brochaTexto = New-Object System.Drawing.SolidBrush($colorTexto)
    try {
        Initialize-Grafico $g
        $g.FillRectangle($brochaFondo, 0, 0, $Ancho, $Alto)
        $margen = [single]($Alto * 0.15)
        $lado = [single]($Alto - 2 * $margen)
        $cuadro = New-RectanguloRedondeado -X $margen -Y $margen -Ancho $lado -Alto $lado -Radio ($lado * 0.2)
        try {
            if ($null -ne $ImagenOrigen) {
                $blanco = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
                try { $g.FillPath($blanco, $cuadro) } finally { $blanco.Dispose() }
                $pad = [single]($lado * 0.08)
                Invoke-DibujoAjustado -Grafico $g -Imagen $ImagenOrigen -X ($margen + $pad) -Y ($margen + $pad) -Ancho ($lado - 2 * $pad) -Alto ($lado - 2 * $pad)
            }
            else {
                $g.FillPath($brochaTexto, $cuadro)
                $fuenteIni = New-Object System.Drawing.Font('Segoe UI', [single]($lado * 0.42), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
                $fmt = New-Object System.Drawing.StringFormat
                $fmt.Alignment = [System.Drawing.StringAlignment]::Center
                $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
                try { $g.DrawString($Iniciales, $fuenteIni, $brochaFondo, (New-Object System.Drawing.RectangleF($margen, $margen, $lado, $lado)), $fmt) }
                finally { $fmt.Dispose(); $fuenteIni.Dispose() }
            }
        }
        finally { $cuadro.Dispose() }

        $fuente = New-Object System.Drawing.Font('Segoe UI', [single]($Alto * 0.30), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $fmtTexto = New-Object System.Drawing.StringFormat
        $fmtTexto.LineAlignment = [System.Drawing.StringAlignment]::Center
        $fmtTexto.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
        $fmtTexto.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
        try {
            $xTexto = [single]($margen * 2 + $lado)
            $g.DrawString($Empresa, $fuente, $brochaTexto, (New-Object System.Drawing.RectangleF($xTexto, 0, ($Ancho - $xTexto - $margen), $Alto)), $fmtTexto)
        }
        finally { $fmtTexto.Dispose(); $fuente.Dispose() }

        $bmp.Save($Ruta, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $brochaTexto.Dispose(); $brochaFondo.Dispose(); $g.Dispose(); $bmp.Dispose()
    }
}
function Get-IconoDeExe {
    # Extrae el ícono de un EXE en alta resolución (256 px). Si no se puede, usa el ícono asociado (32 px).
    param([string]$Ruta)
    try {
        if (-not ('FabricaIconos' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FabricaIconos
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern uint PrivateExtractIcons(string szFileName, int nIconIndex, int cxIcon, int cyIcon, IntPtr[] phicon, uint[] piconid, uint nIcons, uint flags);
    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
'@
        }
        $handles = New-Object IntPtr[] 1
        $ids = New-Object UInt32[] 1
        $n = [FabricaIconos]::PrivateExtractIcons($Ruta, 0, 256, 256, $handles, $ids, 1, 0)
        if ($n -gt 0 -and $n -ne [UInt32]::MaxValue -and $handles[0] -ne [IntPtr]::Zero) {
            try {
                $icono = [System.Drawing.Icon]::FromHandle($handles[0])
                $bmp = $icono.ToBitmap()
                $icono.Dispose()
                return $bmp
            }
            finally { [void][FabricaIconos]::DestroyIcon($handles[0]) }
        }
    }
    catch { }
    try {
        $icono = [System.Drawing.Icon]::ExtractAssociatedIcon($Ruta)
        if ($null -ne $icono) { $bmp = $icono.ToBitmap(); $icono.Dispose(); return $bmp }
    }
    catch { }
    return $null
}
#endregion

#region ===================== Paquete: recursos, logo, procesos y mensaje =====================
function Copy-RecursosCorporativos {
    # Copia Assets, Config y Strings de la plantilla corporativa sobre un paquete.
    param([Parameter(Mandatory)][string]$RutaPaquete)
    foreach ($sub in @('Assets', 'Config', 'Strings')) {
        $origen = Join-Path $script:Recursos $sub
        if (Test-Path -LiteralPath $origen) {
            $destino = Join-Path $RutaPaquete $sub
            if (-not (Test-Path -LiteralPath $destino)) { New-Item -Path $destino -ItemType Directory -Force | Out-Null }
            Get-ChildItem -LiteralPath $destino -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $_.IsReadOnly = $false }
            Copy-Item -Path (Join-Path $origen '*') -Destination $destino -Recurse -Force
        }
    }
}

function Set-LogoVentanas {
    # Pone una imagen como logo de las ventanas de UN paquete (Assets cuadrados; el banner se deja corporativo).
    param([Parameter(Mandatory)][string]$RutaPaquete, [Parameter(Mandatory)][string]$RutaImagen)
    $imagen = Get-ImagenSinBloqueo $RutaImagen
    if ($null -eq $imagen) { Write-Registro -Mensaje "No se pudo abrir el logo de la app: $RutaImagen" -Tipo Aviso; return }
    try {
        foreach ($png in @(Get-ChildItem -LiteralPath (Join-Path $RutaPaquete 'Assets') -Filter '*.png' -File -ErrorAction SilentlyContinue)) {
            $actual = Get-ImagenSinBloqueo $png.FullName
            if ($null -eq $actual) { continue }
            $w = $actual.Width; $h = $actual.Height
            $actual.Dispose()
            if (($w / [double]$h) -lt 3) {
                $png.IsReadOnly = $false
                New-IconoPng -Ruta $png.FullName -Ancho $w -Alto $h -ImagenOrigen $imagen
            }
        }
    }
    finally { $imagen.Dispose() }
}

function Get-ListaProcesos {
    # Interpreta "proceso" o "proceso=Nombre visible" separados por coma. Devuelve una lista de @{ Nombre; Descripcion }.
    param([string]$Procesos, [string]$Descripcion)
    $lista = @()
    if ([string]::IsNullOrWhiteSpace($Procesos)) { return , $lista }
    $partes = @(($Procesos -split '[,;]') | Where-Object { $_.Trim() })
    foreach ($parte in $partes) {
        $nombre = $parte.Trim()
        $visible = ''
        if ($nombre -match '^\s*([^=]+?)\s*=\s*(.+?)\s*$') { $nombre = $Matches[1]; $visible = $Matches[2] }
        $nombre = $nombre -replace '\.exe$', ''
        if (-not $nombre) { continue }
        if (-not $visible) { if ($partes.Count -eq 1 -and $Descripcion) { $visible = $Descripcion } else { $visible = $nombre } }
        $lista += @{ Nombre = $nombre; Descripcion = $visible }
    }
    return , $lista
}

function Get-LiteralProcesos {
    # Devuelve @(@{ Name = 'x'; Description = 'Nombre visible' }, ...) en una sola línea, o @() si no hay procesos.
    param([string]$Procesos, [string]$Descripcion)
    $items = @()
    foreach ($p in (Get-ListaProcesos -Procesos $Procesos -Descripcion $Descripcion)) {
        $items += ('@{ Name = ' + (ConvertTo-TextoPS $p.Nombre) + '; Description = ' + (ConvertTo-TextoPS $p.Descripcion) + ' }')
    }
    if ($items.Count -eq 0) { return '@()' }
    return '@(' + ($items -join ', ') + ')'
}

function Set-TextosVentana {
    # Personaliza los textos de la ventana de bienvenida de PSADT en Strings\<idioma>\strings.psd1 de UN paquete.
    # Claves de CloseAppsPrompt (verificadas en los 27 idiomas de PSADT 4.1.8 y 4.2):
    #   Mensaje        -> Fluent.DialogMessage.Install, Fluent.DialogMessageNoProcesses.Install, Classic.CloseAppsMessage.Install
    #   BotonInstalar  -> Fluent.ButtonLeftText.Install, Fluent.ButtonLeftNoProcessesText.Install, Classic.ButtonClose
    #   BotonAplazar   -> Fluent.ButtonRightText, Classic.ButtonDefer
    # Un valor vacío deja el texto estándar de PSADT.
    param(
        [Parameter(Mandatory)][string]$RutaPaquete,
        [string]$Mensaje,
        [string]$BotonInstalar,
        [string]$BotonAplazar
    )
    $valorEntreComillas = '(' + "'(?:[^']|'')*'" + '|"(?:[^"`]|`.)*")'
    $reglas = @()
    if (-not [string]::IsNullOrWhiteSpace($Mensaje)) {
        foreach ($clave in @('DialogMessage', 'DialogMessageNoProcesses', 'CloseAppsMessage')) {
            $reglas += @{ Patron = '(?m)(^[ \t]*' + $clave + '[ \t]*=[ \t]*@\{[ \t]*\r?\n[ \t]*Install[ \t]*=[ \t]*)' + $valorEntreComillas; Valor = $Mensaje }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($BotonInstalar)) {
        foreach ($clave in @('ButtonLeftText', 'ButtonLeftNoProcessesText')) {
            $reglas += @{ Patron = '(?m)(^[ \t]*' + $clave + '[ \t]*=[ \t]*@\{[ \t]*\r?\n[ \t]*Install[ \t]*=[ \t]*)' + $valorEntreComillas; Valor = $BotonInstalar }
        }
        $reglas += @{ Patron = '(?m)(^[ \t]*ButtonClose[ \t]*=[ \t]*)' + $valorEntreComillas; Valor = $BotonInstalar }
    }
    if (-not [string]::IsNullOrWhiteSpace($BotonAplazar)) {
        foreach ($clave in @('ButtonRightText', 'ButtonDefer')) {
            $reglas += @{ Patron = '(?m)(^[ \t]*' + $clave + '[ \t]*=[ \t]*)' + $valorEntreComillas; Valor = $BotonAplazar }
        }
    }
    if ($reglas.Count -eq 0) { return }

    $cambiados = 0
    foreach ($archivo in @(Get-ChildItem -LiteralPath (Join-Path $RutaPaquete 'Strings') -Filter 'strings.psd1' -Recurse -File -ErrorAction SilentlyContinue)) {
        $texto = [System.IO.File]::ReadAllText($archivo.FullName)
        $original = $texto
        foreach ($regla in $reglas) {
            $rx = New-Object System.Text.RegularExpressions.Regex($regla.Patron)
            $literal = (ConvertTo-TextoPS $regla.Valor).Replace('$', '$$')
            $texto = $rx.Replace($texto, ('${1}' + $literal), 1)
        }
        if ($texto -ne $original) {
            $archivo.IsReadOnly = $false
            [System.IO.File]::WriteAllText($archivo.FullName, $texto, $script:Utf8ConBom)
            $cambiados++
        }
    }
    if ($cambiados -eq 0) { Write-Registro -Mensaje 'No se encontraron los textos estándar en Strings; la ventana mostrará los textos de PSADT.' -Tipo Aviso }
    else { Write-Registro -Mensaje "Textos de la ventana personalizados ($cambiados idiomas)." }
}

function Test-ProcesosEnSesion {
    # La plantilla oficial (4.1+) trae AppProcessesToClose en $adtSession.
    param([string]$RutaScript)
    return ([System.IO.File]::ReadAllText($RutaScript) -match '(?m)^\s*AppProcessesToClose\s*=\s*@\(')
}

function Get-LargoMaximoPlantilla {
    # Largo de la ruta relativa más larga que crea New-ADTTemplate dentro de un paquete
    # (la mide en el módulo PSADT cargado; si no puede, usa un valor prudente).
    if ($script:LargoPlantilla -gt 0) { return $script:LargoPlantilla }
    $largo = 110
    try {
        $modulo = Get-Module -Name 'PSAppDeployToolkit' | Select-Object -First 1
        if ($null -ne $modulo) {
            $base = $modulo.ModuleBase.TrimEnd('\')
            $maximo = 0
            foreach ($f in @(Get-ChildItem -LiteralPath $base -Recurse -File -Force -ErrorAction SilentlyContinue)) {
                $relativo = $f.FullName.Length - $base.Length
                if ($relativo -gt $maximo) { $maximo = $relativo }
            }
            if ($maximo -gt 0) { $largo = [Math]::Max(70, 'PSAppDeployToolkit'.Length + $maximo) }
        }
    }
    catch { }
    $script:LargoPlantilla = $largo
    return $largo
}

function Test-LongitudRutaPaquete {
    # Windows PowerShell 5.1 no maneja rutas de más de 259 caracteres: se valida ANTES de crear el paquete
    # para dar un mensaje claro en vez de «No se puede encontrar una parte de la ruta de acceso».
    param([Parameter(Mandatory)][string]$RutaPaquete, [string]$CarpetaOrigen)
    $limite = 259
    $rutaBase = $RutaPaquete.TrimEnd('\')
    $necesario = $rutaBase.Length + 1 + (Get-LargoMaximoPlantilla)
    $causa = 'los archivos de PSADT'
    if ($CarpetaOrigen -and (Test-Path -LiteralPath $CarpetaOrigen -PathType Container)) {
        $base = $CarpetaOrigen.TrimEnd('\')
        $maximo = 0
        foreach ($f in @(Get-ChildItem -LiteralPath $base -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            $relativo = $f.FullName.Length - $base.Length
            if ($relativo -gt $maximo) { $maximo = $relativo }
        }
        $conArchivos = $rutaBase.Length + '\Files'.Length + $maximo
        if ($conArchivos -gt $necesario) { $necesario = $conArchivos; $causa = 'los archivos de la carpeta del instalador' }
    }
    if ($necesario -gt $limite) {
        throw ("La ruta queda demasiado larga para Windows: {0} caracteres al copiar {1} (el máximo es 259).`n`n" +
            "Solución: copia la carpeta de la Fábrica a una ruta corta y fuera de OneDrive, por ejemplo C:\FabricaPaquetes " +
            "(hay que acortar al menos {2} caracteres). También ayuda un nombre de aplicación más corto.`n`nRuta del paquete: {3}") -f $necesario, $causa, ($necesario - $limite), $rutaBase
    }
}

function Start-VistaPrevia {
    # Arma un paquete de prueba con la configuración actual y muestra las ventanas REALES de PSADT. No instala nada.
    param([Parameter(Mandatory)][hashtable]$Datos)

    $cfg = Get-ConfigCorporativa
    $nombre = '_VistaPrevia'
    $ruta = Join-Path $script:Raiz $nombre
    if (Test-Path -LiteralPath $ruta) {
        Get-ChildItem -LiteralPath $ruta -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.PSIsContainer) { $_.IsReadOnly = $false } }
        Remove-Item -LiteralPath $ruta -Recurse -Force
    }
    Write-Registro -Mensaje 'Preparando vista previa de las ventanas (no se instala nada)...'
    Test-LongitudRutaPaquete -RutaPaquete $ruta
    New-ADTTemplate -Destination $script:Raiz -Name $nombre | Out-Null
    Copy-RecursosCorporativos -RutaPaquete $ruta

    $D = @{}
    foreach ($k in $Datos.Keys) { $D[$k] = $Datos[$k] }
    $D.Tipo = 'VISTA'
    if ([string]::IsNullOrWhiteSpace($D.NombreApp))  { $D.NombreApp = 'Aplicación de prueba' }
    if ([string]::IsNullOrWhiteSpace($D.Version))    { $D.Version = '1.0' }
    if ([string]::IsNullOrWhiteSpace($D.Fabricante)) { $D.Fabricante = [string]$cfg.Empresa }
    $usaNotepad = $false
    if ([string]::IsNullOrWhiteSpace($D.ProcesosACerrar)) { $D.ProcesosACerrar = 'notepad'; $usaNotepad = $true }
    $D.ProcesosEnSesion = $false
    $D.RutaRelativaInstalador = ''

    if ($D.LogoApp -and $D.LogoAppEnVentanas -and (Test-Path -LiteralPath $D.LogoApp)) {
        Set-LogoVentanas -RutaPaquete $ruta -RutaImagen $D.LogoApp
    }
    $rutaScript = Join-Path $ruta 'Invoke-AppDeployToolkit.ps1'
    $valores = [ordered]@{
        AppVendor  = $D.Fabricante
        AppName    = $D.NombreApp
        AppVersion = $D.Version
        AppArch    = $D.Arquitectura
        AppLang    = 'ES'
    }
    Update-ScriptDespliegue -RutaScript $rutaScript -ValoresSesion $valores -Fases (New-ContenidoFases -D $D) -SinRequerirAdmin
    Set-TextosVentana -RutaPaquete $ruta -Mensaje $D.MensajeBienvenida -BotonInstalar $D.TextoBotonInstalar -BotonAplazar $D.TextoBotonAplazar

    if ($usaNotepad -and -not (Get-Process -Name 'notepad' -ErrorAction SilentlyContinue)) {
        Start-Process -FilePath 'notepad.exe'
        Start-Sleep -Seconds 1
    }
    Start-Process -FilePath (Join-Path $ruta 'Invoke-AppDeployToolkit.exe') -ArgumentList '-DeploymentType Install -DeployMode Interactive' -WorkingDirectory $ruta
    Write-Registro -Mensaje 'Vista previa abierta. Así verá las ventanas el usuario (con la app abierta).' -Tipo OK
    if ($usaNotepad) { Write-Registro -Mensaje 'Como no indicaste procesos a cerrar, la vista previa usa el Bloc de notas para mostrar la lista de apps.' }
}
#endregion

#region ===================== Módulo PSADT, plantilla corporativa e inicialización =====================
function Find-ManifiestoLocal {
    $rutaLocal = Join-Path $script:Herramientas 'PSAppDeployToolkit'
    if (-not (Test-Path -LiteralPath $rutaLocal)) { return $null }
    Get-ChildItem -LiteralPath $rutaLocal -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
    return Get-ChildItem -LiteralPath $rutaLocal -Filter 'PSAppDeployToolkit.psd1' -Recurse -File |
        Sort-Object -Property @{ Expression = {
                $v = $null
                if ([version]::TryParse($_.Directory.Name, [ref]$v)) { $v } else { [version]'0.0' }
            } } -Descending |
        Select-Object -First 1
}

function Import-ModuloPSADT {
    if (Get-Command -Name 'New-ADTTemplate' -ErrorAction SilentlyContinue) {
        $yaCargado = Get-Module -Name 'PSAppDeployToolkit' | Select-Object -First 1
        if ($null -ne $yaCargado) { $script:VersionPSADT = $yaCargado.Version }
        return
    }

    $manifiesto = Find-ManifiestoLocal
    if ($null -ne $manifiesto) {
        Import-Module -Name $manifiesto.FullName -Force
    }
    else {
        $modulo = Get-Module -ListAvailable -Name 'PSAppDeployToolkit' | Sort-Object -Property Version -Descending | Select-Object -First 1
        if ($null -ne $modulo) {
            Import-Module -Name $modulo.Path -Force
        }
        else {
            Write-Registro -Mensaje 'No se encontró PSAppDeployToolkit. Intentando descargarlo de la PowerShell Gallery (una sola vez)...'
            try {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                $nuget = Get-PackageProvider -ListAvailable -Name 'NuGet' -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' }
                if (-not $nuget) {
                    Install-PackageProvider -Name 'NuGet' -MinimumVersion '2.8.5.201' -Scope CurrentUser -Force | Out-Null
                }
                Save-Module -Name 'PSAppDeployToolkit' -Path $script:Herramientas -Repository 'PSGallery' -Force
            }
            catch {
                throw ("No se pudo conseguir el módulo PSAppDeployToolkit automáticamente (sin Internet o bloqueado por proxy).`n`n" +
                    "Solución: en un equipo con Internet ejecuta`n  Save-Module -Name PSAppDeployToolkit -Path C:\Temp\PSADT`n" +
                    "y copia la carpeta C:\Temp\PSADT\PSAppDeployToolkit dentro de:`n  $($script:Herramientas)`n`nDetalle: $($_.Exception.Message)")
            }
            $manifiesto = Find-ManifiestoLocal
            if ($null -eq $manifiesto) { throw 'La descarga terminó pero no se encontró PSAppDeployToolkit.psd1 en Herramientas.' }
            Import-Module -Name $manifiesto.FullName -Force
            Write-Registro -Mensaje 'Módulo PSADT descargado en Herramientas: ya queda incluido para la próxima vez.' -Tipo OK
        }
    }

    $cargado = Get-Module -Name 'PSAppDeployToolkit' | Select-Object -First 1
    $script:VersionPSADT = $cargado.Version
    Write-Registro -Mensaje "PSAppDeployToolkit cargado: versión $($cargado.Version)"
    if ($cargado.Version -lt [version]'4.0.0') { throw 'Esta herramienta requiere PSAppDeployToolkit 4.x.' }
    if ($cargado.Version -lt [version]'4.1.0') {
        Write-Registro -Mensaje 'Se recomienda PSADT 4.1 o superior (muestra avisos al usuario desde Intune sin ServiceUI).' -Tipo Aviso
    }
}

function Update-PlantillaCorporativa {
    # Aplica la configuración corporativa a Recursos\Config\config.psd1 y regenera logo y banner en Recursos\Assets.
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Cfg,
        [switch]$NoImportarLogo,
        [switch]$NoImportarBanner
    )

    # --- config.psd1 ---
    $rutaConfig = Join-Path $script:Recursos 'Config\config.psd1'
    if (Test-Path -LiteralPath $rutaConfig) {
        $texto = [System.IO.File]::ReadAllText($rutaConfig)
        $texto = Set-ValorConfig -Texto $texto -Clave 'CompanyName' -ValorLiteral (ConvertTo-TextoPS $Cfg.Empresa) -SoloEntreComillas
        $texto = Set-ValorConfig -Texto $texto -Clave 'DialogStyle' -ValorLiteral (ConvertTo-TextoPS $Cfg.EstiloDialogo) -SoloEntreComillas
        $hex = ([string]$Cfg.ColorAcento).Trim().TrimStart('#')
        if ($hex -match '^[0-9A-Fa-f]{6}$') {
            $texto = Set-ValorConfig -Texto $texto -Clave 'FluentAccentColor' -ValorLiteral ('0xFF' + $hex.ToUpper())
        }
        $idioma = switch ([string]$Cfg.Idioma) { 'es' { "'es'" } 'en' { "'en'" } default { '$null' } }
        $texto = Set-ValorConfig -Texto $texto -Clave 'LanguageOverride' -ValorLiteral $idioma
        if (-not [string]::IsNullOrWhiteSpace([string]$Cfg.RutaLogs)) {
            $texto = Set-ValorConfig -Texto $texto -Clave 'LogPath' -ValorLiteral (ConvertTo-TextoPS $Cfg.RutaLogs) -SoloEntreComillas
        }
        [System.IO.File]::WriteAllText($rutaConfig, $texto, $script:Utf8ConBom)
    }
    else {
        Write-Registro -Mensaje "No existe $rutaConfig; no se aplicó la configuración." -Tipo Aviso
    }

    # --- Logo y banner ---
    # Cada imagen generada se registra con su huella (SHA256) en Corporativo.json. Si al guardar una imagen
    # ya no coincide, es porque alguien la reemplazó a mano en Recursos\Assets: se respeta y se toma como
    # logo/banner propio en lugar de sobrescribirla.
    $carpetaAssets = Join-Path $script:Recursos 'Assets'
    if (-not (Test-Path -LiteralPath $carpetaAssets)) { return }

    $huellas = @{}
    foreach ($par in (([string]$Cfg.HuellasAssets) -split ';')) {
        if ($par -match '^(.+?)=([0-9A-Fa-f]{64})$') { $huellas[$Matches[1]] = $Matches[2] }
    }

    foreach ($png in @(Get-ChildItem -LiteralPath $carpetaAssets -Filter '*.png' -File)) {
        if (-not $huellas.ContainsKey($png.Name)) { continue }
        $actualHuella = (Get-FileHash -LiteralPath $png.FullName -Algorithm SHA256).Hash
        if ($actualHuella -eq $huellas[$png.Name]) { continue }
        $tmp = Get-ImagenSinBloqueo $png.FullName
        if ($null -eq $tmp) { continue }
        $esBanner = (($tmp.Width / [double]$tmp.Height) -ge 3)
        $tmp.Dispose()
        if ($esBanner -and -not $NoImportarBanner) {
            Copy-Item -LiteralPath $png.FullName -Destination (Join-Path $script:Recursos 'BannerEmpresa.png') -Force
            $Cfg.Banner = 'BannerEmpresa.png'
            Write-Registro -Mensaje "Se detectó un banner puesto a mano ($($png.Name)): se conserva como banner propio." -Tipo Aviso
        }
        elseif (-not $esBanner -and -not $NoImportarLogo) {
            Get-ChildItem -LiteralPath $script:Recursos -Filter 'LogoEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            Copy-Item -LiteralPath $png.FullName -Destination (Join-Path $script:Recursos 'LogoEmpresa.png') -Force
            $Cfg.Logo = 'LogoEmpresa.png'
            Write-Registro -Mensaje "Se detectó un logo puesto a mano ($($png.Name)): se conserva como logo de la empresa." -Tipo Aviso
        }
    }

    $imagenLogo = $null
    $imagenBanner = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$Cfg.Logo)) {
        $imagenLogo = Get-ImagenSinBloqueo (Join-Path $script:Recursos ([string]$Cfg.Logo))
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Cfg.Banner)) {
        $imagenBanner = Get-ImagenSinBloqueo (Join-Path $script:Recursos ([string]$Cfg.Banner))
    }
    $iniciales = Get-Iniciales ([string]$Cfg.Empresa)
    $nuevasHuellas = @()
    try {
        foreach ($png in @(Get-ChildItem -LiteralPath $carpetaAssets -Filter '*.png' -File)) {
            $actual = Get-ImagenSinBloqueo $png.FullName
            if ($null -eq $actual) { continue }
            $w = $actual.Width; $h = $actual.Height
            $actual.Dispose()
            $png.IsReadOnly = $false
            if (($w / [double]$h) -ge 3) {
                if ($null -ne $imagenBanner) {
                    New-BannerDesdeImagen -Ruta $png.FullName -Ancho $w -Alto $h -Imagen $imagenBanner
                }
                else {
                    New-BannerPng -Ruta $png.FullName -Ancho $w -Alto $h -ColorHex $Cfg.ColorAcento -Empresa $Cfg.Empresa -Iniciales $iniciales -ImagenOrigen $imagenLogo
                }
            }
            else {
                New-IconoPng -Ruta $png.FullName -Ancho $w -Alto $h -ColorHex $Cfg.ColorAcento -Iniciales $iniciales -ImagenOrigen $imagenLogo
            }
            $nuevasHuellas += ($png.Name + '=' + (Get-FileHash -LiteralPath $png.FullName -Algorithm SHA256).Hash)
        }
    }
    finally {
        if ($null -ne $imagenLogo) { $imagenLogo.Dispose() }
        if ($null -ne $imagenBanner) { $imagenBanner.Dispose() }
    }
    $Cfg.HuellasAssets = ($nuevasHuellas -join ';')
    Save-ConfigCorporativa -Cfg $Cfg
}

function New-BannerDesdeImagen {
    # Ajusta un banner propio al tamaño que usa PSADT. Si la proporción es parecida (±15 %), lo llena
    # completo recortando el centro; si no, lo centra completo sobre fondo blanco sin deformarlo.
    param(
        [Parameter(Mandatory)][string]$Ruta,
        [int]$Ancho = 900,
        [int]$Alto = 125,
        [Parameter(Mandatory)][System.Drawing.Image]$Imagen
    )
    $bmp = New-Object System.Drawing.Bitmap($Ancho, $Alto, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        Initialize-Grafico $g
        $g.Clear([System.Drawing.Color]::White)
        $propDestino = $Ancho / [double]$Alto
        $propOrigen = $Imagen.Width / [double]$Imagen.Height
        if ([Math]::Abs($propOrigen - $propDestino) / $propDestino -le 0.15) {
            $escala = [Math]::Max($Ancho / [double]$Imagen.Width, $Alto / [double]$Imagen.Height)
            $w = [single]($Imagen.Width * $escala)
            $h = [single]($Imagen.Height * $escala)
            $g.DrawImage($Imagen, [single](($Ancho - $w) / 2), [single](($Alto - $h) / 2), $w, $h)
        }
        else {
            Invoke-DibujoAjustado -Grafico $g -Imagen $Imagen -X 0 -Y 0 -Ancho $Ancho -Alto $Alto
        }
        $bmp.Save($Ruta, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $g.Dispose()
        $bmp.Dispose()
    }
}

function Initialize-Fabrica {
    foreach ($carpeta in @($script:Herramientas, $script:CarpetaPaquetes, $script:CarpetaSalida)) {
        if (-not (Test-Path -LiteralPath $carpeta)) { New-Item -Path $carpeta -ItemType Directory -Force | Out-Null }
    }

    Import-ModuloPSADT

    if ($script:Raiz -match '\\OneDrive' -or $script:Raiz.Length -gt 80) {
        Write-Registro -Mensaje "La Fábrica está en una ruta larga o dentro de OneDrive ($($script:Raiz)). Si sale «ruta demasiado larga» o archivos bloqueados, cópiala a una carpeta corta como C:\FabricaPaquetes." -Tipo Aviso
    }

    if (-not (Test-Path -LiteralPath (Join-Path $script:Recursos 'Config\config.psd1'))) {
        Write-Registro -Mensaje 'Primera ejecución: creando la plantilla corporativa automáticamente...'
        $temporal = Join-Path $env:TEMP ('FabricaPSADT_' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $temporal -ItemType Directory -Force | Out-Null
        try {
            New-ADTTemplate -Destination $temporal -Name 'Base' | Out-Null
            if (-not (Test-Path -LiteralPath $script:Recursos)) { New-Item -Path $script:Recursos -ItemType Directory -Force | Out-Null }
            foreach ($sub in @('Assets', 'Config', 'Strings')) {
                $origen = Join-Path $temporal "Base\$sub"
                if (Test-Path -LiteralPath $origen) {
                    Copy-Item -LiteralPath $origen -Destination $script:Recursos -Recurse -Force
                }
            }
            Get-ChildItem -LiteralPath $script:Recursos -Recurse -File | ForEach-Object { $_.IsReadOnly = $false }
        }
        finally {
            Get-ChildItem -LiteralPath $temporal -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.PSIsContainer) { $_.IsReadOnly = $false } }
            Remove-Item -LiteralPath $temporal -Recurse -Force -ErrorAction SilentlyContinue
        }
        $cfg = Get-ConfigCorporativa
        Save-ConfigCorporativa -Cfg $cfg
        Update-PlantillaCorporativa -Cfg $cfg
        Write-Registro -Mensaje "Plantilla corporativa lista (empresa: $($cfg.Empresa)). Puedes cambiarla en la pestaña 'Plantilla corporativa'." -Tipo OK
    }
    elseif (-not (Test-Path -LiteralPath $script:RutaConfigJson)) {
        Save-ConfigCorporativa -Cfg (Get-ConfigCorporativa)
    }

    if (-not (Test-Path -LiteralPath (Join-Path $script:Herramientas 'IntuneWinAppUtil.exe'))) {
        Write-Registro -Mensaje "No está IntuneWinAppUtil.exe en '$($script:Herramientas)'. Podrás crear paquetes, pero no el .intunewin." -Tipo Aviso
    }
}
#endregion

#region ===================== Lectura del instalador =====================
function Get-PropiedadesMsi {
    param([Parameter(Mandatory)][string]$Ruta)
    $resultado  = @{}
    $instalador = $null
    $baseDatos  = $null
    try {
        $instalador = New-Object -ComObject WindowsInstaller.Installer
        $baseDatos  = $instalador.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $instalador, @($Ruta, 0))
        foreach ($propiedad in @('ProductName', 'ProductVersion', 'Manufacturer', 'ProductCode')) {
            $vista = $baseDatos.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $baseDatos, @("SELECT Value FROM Property WHERE Property='$propiedad'"))
            $vista.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $vista, $null) | Out-Null
            $registro = $vista.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $vista, $null)
            if ($null -ne $registro) {
                $resultado[$propiedad] = [string]$registro.GetType().InvokeMember('StringData', 'GetProperty', $null, $registro, 1)
            }
            $vista.GetType().InvokeMember('Close', 'InvokeMethod', $null, $vista, $null) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vista) | Out-Null
        }
    }
    finally {
        if ($null -ne $baseDatos)  { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($baseDatos)  | Out-Null }
        if ($null -ne $instalador) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($instalador) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    return $resultado
}

function Get-InfoInstalador {
    param([Parameter(Mandatory)][string]$Ruta)

    if (-not (Test-Path -LiteralPath $Ruta -PathType Leaf)) { throw "No existe el instalador: $Ruta" }

    $info = @{ Tipo = ''; Fabricante = ''; Nombre = ''; Version = ''; CodigoProducto = ''; Desinstalador = '' }
    $extension = [System.IO.Path]::GetExtension($Ruta).ToLowerInvariant()

    switch ($extension) {
        '.msi' {
            $info.Tipo = 'MSI'
            $p = Get-PropiedadesMsi -Ruta $Ruta
            if ($p.ContainsKey('Manufacturer'))   { $info.Fabricante     = $p['Manufacturer'] }
            if ($p.ContainsKey('ProductName'))    { $info.Nombre         = $p['ProductName'] }
            if ($p.ContainsKey('ProductVersion')) { $info.Version        = $p['ProductVersion'] }
            if ($p.ContainsKey('ProductCode'))    { $info.CodigoProducto = $p['ProductCode'] }
        }
        '.exe' {
            $info.Tipo = 'EXE'
            $vi = (Get-Item -LiteralPath $Ruta).VersionInfo
            $info.Fabricante = [string]$vi.CompanyName
            $info.Nombre     = [string]$vi.ProductName
            if ([string]::IsNullOrWhiteSpace($info.Nombre)) { $info.Nombre = [string]$vi.FileDescription }
            $info.Version    = ([string]$vi.ProductVersion) -replace ',\s*', '.'
        }
        '.ps1' {
            $info.Tipo = 'PS1'
            $carpeta = Split-Path -Path $Ruta -Parent
            $base = [System.IO.Path]::GetFileNameWithoutExtension($Ruta)
            if ($base -match '^(install|instalar|instalacion|setup|deploy|despliegue|main|run|ejecutar)$') {
                $info.Nombre = Split-Path -Path $carpeta -Leaf
            }
            else { $info.Nombre = $base }
            $info.Version = '1.0.0'
            $candidato = Get-ChildItem -LiteralPath $carpeta -Filter '*.ps1' -File |
                Where-Object { $_.FullName -ne $Ruta -and $_.BaseName -match 'desinst|uninst|remov|quitar|borrar' } |
                Select-Object -First 1
            if ($null -ne $candidato) { $info.Desinstalador = $candidato.FullName }
        }
        default { throw 'Solo se admiten instaladores .msi, .exe o scripts .ps1.' }
    }

    $info.Fabricante = ([string]$info.Fabricante).Trim()
    $info.Nombre     = ([string]$info.Nombre).Trim()
    $info.Version    = ([string]$info.Version).Trim()
    return $info
}
#endregion

#region ===================== Generación del script de despliegue =====================
function Set-ValorSesion {
    # Cambia el valor de una clave del hashtable $adtSession (ej. AppVendor = '').
    param([string]$Texto, [string]$Clave, [string]$Valor)
    $patron = "(?m)^(\s*$Clave\s*=\s*)'[^']*'"
    $rx = New-Object System.Text.RegularExpressions.Regex($patron)
    if (-not $rx.IsMatch($Texto)) {
        Write-Registro -Mensaje "No se encontró '$Clave' en la plantilla; se deja como está." -Tipo Aviso
        return $Texto
    }
    $reemplazo = '${1}' + "'" + ($Valor -replace "'", "''").Replace('$', '$$') + "'"
    return $rx.Replace($Texto, $reemplazo, 1)
}

function New-ContenidoFases {
    # Devuelve un hashtable Fase -> líneas de código (sin indentar).
    param([Parameter(Mandatory)][hashtable]$D)

    # --- Procesos a cerrar ---
    $bloqueProcesos = @('## Procesos que se deben cerrar antes de continuar')
    if ($D.ProcesosEnSesion) {
        # Definidos en $adtSession.AppProcessesToClose (comportamiento oficial de PSADT: si la app
        # está cerrada, la instalación se hace en silencio).
        $bloqueProcesos += '$procesosACerrar = $adtSession.AppProcessesToClose'
    }
    else {
        $bloqueProcesos += ('$procesosACerrar = ' + (Get-LiteralProcesos -Procesos $D.ProcesosACerrar -Descripcion $D.NombreApp))
    }

    # --- Rutas y comandos ---
    $rutaEnFiles = '(Join-Path -Path $adtSession.DirFiles -ChildPath ' + (ConvertTo-TextoPS $D.RutaRelativaInstalador) + ')'
    $instalar = @(); $desinstalar = @(); $reparar = @(); $marcaPost = @(); $marcaQuitar = @()

    switch ($D.Tipo) {
        'MSI' {
            $l = 'Start-ADTMsiProcess -Action Install -FilePath ' + $rutaEnFiles
            if (-not [string]::IsNullOrWhiteSpace($D.ArgumentosInstalacion)) { $l += ' -AdditionalArgumentList ' + (ConvertTo-TextoPS $D.ArgumentosInstalacion) }
            $instalar    = @($l)
            $desinstalar = @('Start-ADTMsiProcess -Action Uninstall -FilePath ' + $rutaEnFiles)
            $reparar     = @('Start-ADTMsiProcess -Action Repair -FilePath ' + $rutaEnFiles)
        }
        'EXE' {
            $l = 'Start-ADTProcess -FilePath ' + $rutaEnFiles
            if (-not [string]::IsNullOrWhiteSpace($D.ArgumentosInstalacion)) { $l += ' -ArgumentList ' + (ConvertTo-TextoPS $D.ArgumentosInstalacion) }
            $instalar = @($l)
            $ld = 'Start-ADTProcess -FilePath ' + (ConvertTo-TextoPSExpandible $D.DesinstaladorExe)
            if (-not [string]::IsNullOrWhiteSpace($D.ArgumentosDesinstalacion)) { $ld += ' -ArgumentList ' + (ConvertTo-TextoPS $D.ArgumentosDesinstalacion) }
            $desinstalar = @($ld)
            $reparar     = @($l)
        }
        'PS1' {
            $ps = '$powershell = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"'
            $instalar = @(
                $ps,
                ('$scriptInstalacion = Join-Path -Path $adtSession.DirFiles -ChildPath ' + (ConvertTo-TextoPS $D.RutaRelativaInstalador)),
                '$argumentos = ''-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'' + $scriptInstalacion + ''"'''
            )
            if (-not [string]::IsNullOrWhiteSpace($D.ArgumentosInstalacion)) { $instalar += ('$argumentos += '' '' + ' + (ConvertTo-TextoPS $D.ArgumentosInstalacion)) }
            $instalar += 'Start-ADTProcess -FilePath $powershell -ArgumentList $argumentos -CreateNoWindow'
            $reparar = $instalar

            if (-not [string]::IsNullOrWhiteSpace($D.RutaRelativaDesinstalador)) {
                $desinstalar = @(
                    $ps,
                    ('$scriptDesinstalacion = Join-Path -Path $adtSession.DirFiles -ChildPath ' + (ConvertTo-TextoPS $D.RutaRelativaDesinstalador)),
                    '$argumentos = ''-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'' + $scriptDesinstalacion + ''"'''
                )
                if (-not [string]::IsNullOrWhiteSpace($D.ArgumentosDesinstalacion)) { $desinstalar += ('$argumentos += '' '' + ' + (ConvertTo-TextoPS $D.ArgumentosDesinstalacion)) }
                $desinstalar += 'Start-ADTProcess -FilePath $powershell -ArgumentList $argumentos -CreateNoWindow'
            }
            else {
                $desinstalar = @('## Sin script de desinstalación: solo se elimina la marca de registro (Post-Uninstall).')
            }

            $clave = ConvertTo-TextoPS $D.ClaveMarca
            $marcaPost = @(
                '',
                '## Marca de instalación (la usa la detección de Intune)',
                ('Set-ADTRegistryKey -LiteralPath ' + $clave + ' -Name ''Version'' -Value $adtSession.AppVersion -Type String'),
                ('Set-ADTRegistryKey -LiteralPath ' + $clave + ' -Name ''FechaInstalacion'' -Value (Get-Date -Format ''yyyy-MM-dd HH:mm:ss'') -Type String')
            )
            $marcaQuitar = @(
                '## Quitar la marca de instalación',
                ('Remove-ADTRegistryKey -LiteralPath ' + $clave + ' -Recurse')
            )
        }
        'VISTA' {
            # Solo para la vista previa: simula la instalación sin instalar nada.
            $instalar    = @('## VISTA PREVIA: no se instala nada', 'Start-Sleep -Seconds 5')
            $desinstalar = @('Start-Sleep -Seconds 2')
            $reparar     = @('Start-Sleep -Seconds 2')
        }
    }

    # --- Ventana de bienvenida ---
    # Combinaciones de parámetros validadas contra los conjuntos de parámetros de
    # Show-ADTInstallationWelcome en PSADT 4.1.8 y 4.2.
    $cuentaSegundos = [int]$D.CuentaRegresivaMin * 60
    $preInstalar = @() + $bloqueProcesos + @(
        '',
        '## Ventana de bienvenida: cierre de aplicaciones, aplazamientos y espacio en disco',
        '$saiwParams = @{',
        '    CheckDiskSpace = $true',
        '}'
    )
    if ($D.PermitirAplazar) {
        $preInstalar += '$saiwParams.Add(''AllowDefer'', $true)'
        $preInstalar += ('$saiwParams.Add(''DeferTimes'', ' + [int]$D.VecesAplazar + ')')
        $preInstalar += '$saiwParams.Add(''PersistPrompt'', $true)'
        if ([int]$D.DiasLimite -gt 0) {
            $preInstalar += ('$saiwParams.Add(''DeferDays'', ' + [int]$D.DiasLimite + ')   # fecha límite: N días desde el primer intento')
        }
    }
    # El texto del usuario reemplaza el texto estándar de la ventana (ver Set-TextosVentana).
    $preInstalar += @(
        'if ($procesosACerrar.Count -gt 0)',
        '{',
        '    $saiwParams.Add(''CloseProcesses'', $procesosACerrar)'
    )
    if ($cuentaSegundos -gt 0) {
        $preInstalar += ('    $saiwParams.Add(''ForceCloseProcessesCountdown'', ' + $cuentaSegundos + ')   # cuenta regresiva para cerrar las apps e instalar')
    }
    $preInstalar += '}'
    if ($cuentaSegundos -gt 0 -and $D.PermitirAplazar) {
        $preInstalar += @(
            'else',
            '{',
            ('    $saiwParams.Add(''ForceCountdown'', ' + $cuentaSegundos + ')   # cuenta regresiva para iniciar la instalación'),
            '}'
        )
    }
    $preInstalar += @(
        'Show-ADTInstallationWelcome @saiwParams',
        '',
        '## Barra de progreso',
        'Show-ADTInstallationProgress'
    )

    $preOtros = @() + $bloqueProcesos + @(
        'if ($procesosACerrar.Count -gt 0)',
        '{',
        '    Show-ADTInstallationWelcome -CloseProcesses $procesosACerrar -CloseProcessesCountdown 60',
        '}',
        'Show-ADTInstallationProgress'
    )

    $postInstalar = @('## Tareas posteriores a la instalación') + $marcaPost
    if ($D.Tipo -eq 'VISTA' -and [string]::IsNullOrWhiteSpace($D.MensajeFinal)) {
        $postInstalar += 'Show-ADTInstallationPrompt -Message ''Vista previa terminada. No se instaló nada.'' -ButtonRightText ''Aceptar'' -Icon Information'
    }
    if (-not [string]::IsNullOrWhiteSpace($D.MensajeFinal)) {
        $postInstalar += @('', '## Mensaje final al usuario', ('Show-ADTInstallationPrompt -Message ' + (ConvertTo-TextoPS $D.MensajeFinal) + ' -ButtonRightText ''Aceptar'' -Icon Information -NoWait'))
    }

    $postDesinstalar = @('## Tareas posteriores a la desinstalación') + $marcaQuitar
    $postReparar     = @('## Tareas posteriores a la reparación') + $marcaPost

    return @{
        'Pre-Install'    = $preInstalar
        'Install'        = $instalar
        'Post-Install'   = $postInstalar
        'Pre-Uninstall'  = $preOtros
        'Uninstall'      = $desinstalar
        'Post-Uninstall' = $postDesinstalar
        'Pre-Repair'     = $preOtros
        'Repair'         = $reparar
        'Post-Repair'    = $postReparar
    }
}

function ConvertTo-BloqueFunciones {
    # Formato PSADT 4.0 / 4.1: funciones Install/Uninstall/Repair-ADTDeployment.
    param([hashtable]$Fases)
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add('##================================================')
    $out.Add('## MARK: Funciones generadas por la Fabrica de Paquetes (' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + ')')
    $out.Add('##================================================')
    $out.Add('')
    foreach ($accion in @('Install', 'Uninstall', 'Repair')) {
        $out.Add("function $accion-ADTDeployment")
        $out.Add('{')
        $out.Add('    [CmdletBinding()]')
        $out.Add('    param')
        $out.Add('    (')
        $out.Add('    )')
        $out.Add('')
        $out.Add("    ## MARK: Pre-$accion")
        $out.Add('    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"')
        $out.AddRange([string[]](ConvertTo-Indentado -Lineas $Fases["Pre-$accion"]))
        $out.Add('')
        $out.Add("    ## MARK: $accion")
        $out.Add('    $adtSession.InstallPhase = $adtSession.DeploymentType')
        $out.AddRange([string[]](ConvertTo-Indentado -Lineas $Fases[$accion]))
        $out.Add('')
        $out.Add("    ## MARK: Post-$accion")
        $out.Add('    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"')
        $out.AddRange([string[]](ConvertTo-Indentado -Lineas $Fases["Post-$accion"]))
        $out.Add('}')
        $out.Add('')
    }
    $out.Add('## MARK: Initialization')
    return , $out.ToArray()
}

function ConvertTo-BloqueVariables {
    # Formato PSADT 4.2: bloques New-Variable -Name <Fase> -Value { ... }.
    param([hashtable]$Fases)
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add('##================================================')
    $out.Add('## Fases generadas por la Fabrica de Paquetes (' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + ')')
    $out.Add('##================================================')
    $out.Add('')
    foreach ($fase in $script:OrdenFases) {
        $out.Add("## MARK: $fase")
        $out.Add("New-Variable -Name $fase -Force -Value {")
        $out.AddRange([string[]](ConvertTo-Indentado -Lineas $Fases[$fase]))
        $out.Add('}')
        $out.Add('')
    }
    $out.Add('## MARK: Initialization')
    return , $out.ToArray()
}

function Update-ScriptDespliegue {
    param(
        [Parameter(Mandatory)][string]$RutaScript,
        [Parameter(Mandatory)][System.Collections.IDictionary]$ValoresSesion,
        [Parameter(Mandatory)][hashtable]$Fases,
        [string]$ProcesosSesionLiteral,
        [switch]$SinRequerirAdmin
    )

    $texto = [System.IO.File]::ReadAllText($RutaScript)
    foreach ($clave in $ValoresSesion.Keys) {
        $texto = Set-ValorSesion -Texto $texto -Clave $clave -Valor ([string]$ValoresSesion[$clave])
    }
    if (-not [string]::IsNullOrWhiteSpace($ProcesosSesionLiteral)) {
        $rxProc = New-Object System.Text.RegularExpressions.Regex('(?m)^(\s*AppProcessesToClose\s*=\s*)@\([^\r\n]*')
        if ($rxProc.IsMatch($texto)) {
            $texto = $rxProc.Replace($texto, ('${1}' + $ProcesosSesionLiteral.Replace('$', '$$')), 1)
        }
        else { throw 'La plantilla no tiene AppProcessesToClose en $adtSession.' }
    }
    if ($SinRequerirAdmin) {
        $rxAdmin = New-Object System.Text.RegularExpressions.Regex('(?m)^(\s*RequireAdmin\s*=\s*)\$true')
        $texto = $rxAdmin.Replace($texto, '${1}$$false', 1)
    }

    $lineas = $texto -split "`r?`n"
    $inicio = -1
    $fin = -1
    $formato = ''

    for ($i = 0; $i -lt $lineas.Count; $i++) {
        if ($lineas[$i] -match '^\s*function\s+Install-ADTDeployment\b') { $inicio = $i; $formato = 'Funciones'; break }
    }
    if ($inicio -lt 0) {
        for ($i = 0; $i -lt $lineas.Count; $i++) {
            if ($lineas[$i] -match '^\s*##\s*MARK:\s*Pre-Install\s*$') { $inicio = $i; $formato = 'Variables'; break }
            if ($lineas[$i] -match '^\s*New-Variable\s+-Name\s+Pre-Install\b') { $inicio = $i; $formato = 'Variables'; break }
        }
    }
    if ($inicio -ge 0) {
        for ($i = $inicio + 1; $i -lt $lineas.Count; $i++) {
            if ($lineas[$i] -match '^\s*\$ErrorActionPreference\s*=') { $fin = $i; break }
        }
    }
    if ($inicio -lt 0 -or $fin -lt 0) {
        throw 'No se reconoció la estructura de la plantilla PSADT de este módulo. No se modificó el script.'
    }

    $region = ($lineas[$inicio..($fin - 1)]) -join "`n"
    if ($formato -eq 'Funciones') {
        if ($region -notmatch 'function\s+Uninstall-ADTDeployment' -or $region -notmatch 'function\s+Repair-ADTDeployment') {
            throw 'La plantilla PSADT no tiene la estructura esperada (Install/Uninstall/Repair-ADTDeployment). No se modificó el script.'
        }
        $bloque = ConvertTo-BloqueFunciones -Fases $Fases
    }
    else {
        if ($region -notmatch 'New-Variable\s+-Name\s+Uninstall\b' -or $region -notmatch 'New-Variable\s+-Name\s+Repair\b') {
            throw 'La plantilla PSADT no tiene la estructura esperada (fases New-Variable). No se modificó el script.'
        }
        $bloque = ConvertTo-BloqueVariables -Fases $Fases
    }

    $nuevo = New-Object System.Collections.Generic.List[string]
    if ($inicio -gt 0) { $nuevo.AddRange([string[]]$lineas[0..($inicio - 1)]) }
    $nuevo.AddRange([string[]]$bloque)
    $nuevo.AddRange([string[]]$lineas[$fin..($lineas.Count - 1)])
    [System.IO.File]::WriteAllText($RutaScript, ($nuevo -join "`r`n"), $script:Utf8ConBom)
    Write-Registro -Mensaje "Plantilla PSADT detectada: formato '$formato'."
}
#endregion

#region ===================== Detección y ficha de Intune =====================
function New-ScriptDeteccion {
    param([Parameter(Mandatory)][hashtable]$D, [Parameter(Mandatory)][string]$Ruta)

    $encabezado = @(
        '# Script de deteccion generado por la Fabrica de Paquetes'
        ('# Aplicacion: ' + $D.NombreApp + ' ' + $D.Version)
        '# Intune: "Ejecutar script como proceso de 32 bits en clientes de 64 bits" = No'
    ) -join "`r`n"

    if ($D.Tipo -eq 'MSI' -and -not [string]::IsNullOrWhiteSpace($D.CodigoProducto)) {
        $cuerpo = @'
$codigoProducto = __CODIGO__
$claves = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$codigoProducto",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$codigoProducto"
)
foreach ($clave in $claves)
{
    if (Test-Path -LiteralPath $clave)
    {
        Write-Output "Detectado: $codigoProducto"
        exit 0
    }
}
exit 1
'@
        $cuerpo = $cuerpo.Replace('__CODIGO__', (ConvertTo-TextoPS $D.CodigoProducto))
    }
    elseif ($D.Tipo -eq 'PS1') {
        $cuerpo = @'
$clave           = __CLAVE__
$versionEsperada = __VERSION__
if (-not (Test-Path -LiteralPath $clave)) { exit 1 }
$valor = [string](Get-ItemProperty -LiteralPath $clave -ErrorAction SilentlyContinue).Version
if ([string]::IsNullOrWhiteSpace($valor)) { exit 1 }
$instalada = $null
$requerida = $null
if ([version]::TryParse($valor, [ref]$instalada) -and [version]::TryParse($versionEsperada, [ref]$requerida))
{
    if ($instalada -ge $requerida) { Write-Output "Detectado: version $valor"; exit 0 }
    exit 1
}
if ($valor -eq $versionEsperada) { Write-Output "Detectado: version $valor"; exit 0 }
exit 1
'@
        $cuerpo = $cuerpo.Replace('__CLAVE__', (ConvertTo-TextoPS $D.ClaveMarca)).Replace('__VERSION__', (ConvertTo-TextoPS $D.Version))
    }
    else {
        $cuerpo = @'
$nombre        = __NOMBRE__
$versionMinima = __VERSION__
$claves = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$apps = Get-ItemProperty -Path $claves -ErrorAction SilentlyContinue |
    Where-Object { $_.PSObject.Properties['DisplayName'] -and ([string]$_.DisplayName) -like "*$nombre*" }
foreach ($app in $apps)
{
    if ([string]::IsNullOrWhiteSpace($versionMinima)) { Write-Output "Detectado: $($app.DisplayName)"; exit 0 }
    if (-not $app.PSObject.Properties['DisplayVersion']) { continue }
    $instalada = $null
    $requerida = $null
    if ([version]::TryParse([string]$app.DisplayVersion, [ref]$instalada) -and [version]::TryParse($versionMinima, [ref]$requerida))
    {
        if ($instalada -ge $requerida) { Write-Output "Detectado: $($app.DisplayName) $($app.DisplayVersion)"; exit 0 }
    }
    elseif ([string]$app.DisplayVersion -eq $versionMinima)
    {
        Write-Output "Detectado: $($app.DisplayName) $($app.DisplayVersion)"; exit 0
    }
}
exit 1
'@
        $cuerpo = $cuerpo.Replace('__NOMBRE__', (ConvertTo-TextoPS $D.NombreDeteccion)).Replace('__VERSION__', (ConvertTo-TextoPS $D.Version))
    }

    [System.IO.File]::WriteAllText($Ruta, ($encabezado + "`r`n" + $cuerpo), $script:Utf8ConBom)
}

function New-FichaIntune {
    param([Parameter(Mandatory)][hashtable]$D, [Parameter(Mandatory)][string]$Ruta, [string]$CodigoAplazamiento)

    switch ($D.Tipo) {
        'MSI' {
            if ($D.CodigoProducto) { $deteccion = "Regla MSI con código de producto $($D.CodigoProducto)   (o script Deteccion.ps1)" }
            else { $deteccion = 'Script personalizado: Deteccion.ps1' }
        }
        'PS1' {
            $claveIntune = $D.ClaveMarca -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'
            $deteccion = "Script personalizado: Deteccion.ps1`r`n  (o regla de Registro: clave $claveIntune  valor 'Version'  comparación de cadena = $($D.Version))"
        }
        default {
            if ([string]::IsNullOrWhiteSpace($D.Version)) { $deteccion = "Script personalizado: Deteccion.ps1  (busca '$($D.NombreDeteccion)' en cualquier versión)" }
            else { $deteccion = "Script personalizado: Deteccion.ps1  (busca '$($D.NombreDeteccion)' versión >= $($D.Version))" }
        }
    }

    if ($CodigoAplazamiento) { $lineaAplazamiento = "  $CodigoAplazamiento  -> Reintentar   (el usuario aplazó la instalación)" }
    else { $lineaAplazamiento = '  (Revisa DeferExitCode en Config\config.psd1 y agrégalo como Reintentar)' }

    if ($D.Arquitectura -eq 'x64') { $arquitectura = '64 bits' } else { $arquitectura = '32 y 64 bits' }

    if ([string]::IsNullOrWhiteSpace($D.ProcesosACerrar)) {
        $experiencia = 'Sin procesos a cerrar: con usuario conectado se muestran bienvenida/aplazamiento y progreso.'
    }
    elseif ($D.ProcesosEnSesion) {
        $experiencia = "Procesos: $($D.ProcesosACerrar). Si están ABIERTOS se muestran las ventanas; si están cerrados se instala en silencio (comportamiento oficial PSADT)."
    }
    else {
        $experiencia = "Procesos: $($D.ProcesosACerrar). Las ventanas se muestran siempre que haya un usuario conectado."
    }

    $ficha = @"
==============================================================
 FICHA DE CONFIGURACIÓN INTUNE - $($D.NombrePaquete)
 Generada: $(Get-Date -Format 'yyyy-MM-dd HH:mm')  por $env:USERNAME
==============================================================

[Información de la aplicación]
  Nombre ............: $($D.NombreApp)
  Publicador ........: $($D.Fabricante)
  Versión ...........: $($D.Version)
  Descripción .......: $($D.Descripcion)
  Archivo ...........: $($D.NombrePaquete).intunewin
  Logo ..............: Logo_Intune.png  (en esta misma carpeta)

[Programa]
  Comando instalar ..: Invoke-AppDeployToolkit.exe -DeploymentType Install
  Comando desinstalar: Invoke-AppDeployToolkit.exe -DeploymentType Uninstall
  Duración permitida : 60 minutos
  Comportamiento ....: Sistema
  Reinicio ..........: Determinar el comportamiento según los códigos de retorno

[Experiencia del usuario]
  $experiencia

[Códigos de retorno]
  0     -> Correcto
  1707  -> Correcto
  3010  -> Reinicio suave
  1641  -> Reinicio forzado
  1618  -> Reintentar
$lineaAplazamiento

[Requisitos]
  Arquitectura ......: $arquitectura
  SO mínimo .........: Windows 10 (ajustar según el estándar)

[Detección]
  $deteccion

[Logs en el equipo]
  Ruta configurada en la plantilla corporativa (por defecto C:\Windows\Logs\Software)

[Pruebas antes de publicar]
  1. psexec -s -i cmd
  2. cd /d "$($D.RutaPaquete)"
  3. Invoke-AppDeployToolkit.exe -DeploymentType Install
  4. powershell -ExecutionPolicy Bypass -File "$($D.RutaSalida)\Deteccion.ps1"    (debe decir Detectado)
  5. Invoke-AppDeployToolkit.exe -DeploymentType Uninstall
  6. Repetir el paso 4 (no debe detectar nada)
"@
    [System.IO.File]::WriteAllText($Ruta, $ficha, $script:Utf8ConBom)
}
#endregion

#region ===================== Proceso principal =====================
function New-PaqueteCorporativo {
    param([Parameter(Mandatory)][hashtable]$Datos)

    Write-Registro -Mensaje '------ Iniciando creación del paquete ------'
    $cfg = Get-ConfigCorporativa

    # 1. Leer instalador y completar datos
    $info = Get-InfoInstalador -Ruta $Datos.Instalador
    $D = @{}
    foreach ($k in $Datos.Keys) { $D[$k] = $Datos[$k] }
    $D.Tipo           = $info.Tipo
    $D.CodigoProducto = $info.CodigoProducto
    if ([string]::IsNullOrWhiteSpace($D.Fabricante)) { $D.Fabricante = $info.Fabricante }
    if ([string]::IsNullOrWhiteSpace($D.Fabricante)) { $D.Fabricante = [string]$cfg.Empresa }
    if ([string]::IsNullOrWhiteSpace($D.NombreApp))  { $D.NombreApp  = $info.Nombre }
    if ([string]::IsNullOrWhiteSpace($D.Version))    { $D.Version    = $info.Version }
    if ([string]::IsNullOrWhiteSpace($D.Version) -and $D.Tipo -eq 'PS1') {
        # La marca de detección del PS1 la crea el propio paquete: cualquier versión sirve.
        $D.Version = '1.0.0'
        Write-Registro -Mensaje 'Sin versión: se usa 1.0.0 para la marca de detección del script.' -Tipo Aviso
    }
    elseif ([string]::IsNullOrWhiteSpace($D.Version)) {
        $D.Version = ''
        Write-Registro -Mensaje 'Sin versión: la detección buscará la aplicación solo por su nombre, sin comparar versión.' -Tipo Aviso
    }
    if ([string]::IsNullOrWhiteSpace($D.NombreDeteccion)) { $D.NombreDeteccion = $D.NombreApp }
    if ([string]::IsNullOrWhiteSpace($D.Descripcion)) {
        $D.Descripcion = (("Instala $($D.NombreApp) $($D.Version)").Trim() + " de $($D.Fabricante).")
    }
    if ($D.Tipo -eq 'PS1') { $D.CopiarCarpetaCompleta = $true }
    if ($D.Tipo -eq 'EXE' -and [string]::IsNullOrWhiteSpace($D.LogoApp)) {
        $icono = Get-IconoDeExe -Ruta $D.Instalador
        if ($null -ne $icono) {
            $D.LogoApp = Join-Path $env:TEMP ('FabricaIcono_' + [guid]::NewGuid().ToString('N') + '.png')
            try { $icono.Save($D.LogoApp, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $icono.Dispose() }
        }
    }

    $carpetaOrigen = Split-Path -Path $D.Instalador -Parent
    $baseOrigen = $carpetaOrigen.TrimEnd('\') + '\'
    if ($D.CopiarCarpetaCompleta) { $D.RutaRelativaInstalador = $D.Instalador.Substring($baseOrigen.Length) }
    else { $D.RutaRelativaInstalador = [System.IO.Path]::GetFileName($D.Instalador) }

    $D.RutaRelativaDesinstalador = ''
    if ($D.Tipo -eq 'PS1' -and -not [string]::IsNullOrWhiteSpace($D.DesinstaladorExe)) {
        $completo = [System.IO.Path]::GetFullPath($D.DesinstaladorExe)
        if (-not $completo.StartsWith($baseOrigen, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'El script de desinstalación debe estar en la misma carpeta (o una subcarpeta) que el script de instalación.'
        }
        if (-not (Test-Path -LiteralPath $completo -PathType Leaf)) { throw "No existe el script de desinstalación: $completo" }
        $D.RutaRelativaDesinstalador = $completo.Substring($baseOrigen.Length)
    }

    # 2. Validaciones
    $faltantes = @()
    if ([string]::IsNullOrWhiteSpace($D.NombreApp)) { $faltantes += 'Nombre de la aplicación' }
    if ($D.Tipo -eq 'EXE' -and [string]::IsNullOrWhiteSpace($D.DesinstaladorExe)) { $faltantes += 'Desinstalador (obligatorio para EXE)' }
    if ($faltantes.Count -gt 0) { throw ('Faltan datos: ' + ($faltantes -join ', ')) }
    if ($D.Tipo -eq 'EXE' -and [string]::IsNullOrWhiteSpace($D.ArgumentosInstalacion)) {
        Write-Registro -Mensaje 'El EXE no tiene argumentos silenciosos. Si el instalador muestra ventanas, Intune quedará esperando.' -Tipo Aviso
    }
    if ($D.Tipo -eq 'PS1' -and -not $D.RutaRelativaDesinstalador) {
        Write-Registro -Mensaje 'PS1 sin script de desinstalación: al desinstalar solo se quitará la marca de registro.' -Tipo Aviso
    }

    $D.NombrePaquete = Get-NombreLimpio ((@($D.Fabricante, $D.NombreApp, $D.Version, $D.Arquitectura) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '_')
    $D.RutaPaquete   = Join-Path $script:CarpetaPaquetes $D.NombrePaquete
    $baseSalida = $script:CarpetaSalida
    if (-not [string]::IsNullOrWhiteSpace($D.CarpetaDestino)) {
        $baseSalida = [System.IO.Path]::GetFullPath(([string]$D.CarpetaDestino).Trim())
        $paquetesNorm = $script:CarpetaPaquetes.TrimEnd('\') + '\'
        if (($baseSalida.TrimEnd('\') + '\').StartsWith($paquetesNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "La carpeta destino no puede estar dentro de '$($script:CarpetaPaquetes)'. Elige otra carpeta."
        }
        if (-not (Test-Path -LiteralPath $baseSalida -PathType Container)) {
            try { New-Item -Path $baseSalida -ItemType Directory -Force | Out-Null }
            catch { throw "No se pudo crear la carpeta destino '$baseSalida': $($_.Exception.Message)" }
        }
        $prueba = Join-Path $baseSalida ('.fabrica_prueba_' + [guid]::NewGuid().ToString('N'))
        try { [System.IO.File]::WriteAllText($prueba, 'ok'); Remove-Item -LiteralPath $prueba -Force }
        catch { throw "No tienes permiso de escritura en la carpeta destino '$baseSalida'." }
    }
    $D.RutaSalida    = Join-Path $baseSalida $D.NombrePaquete
    $empresaLimpia   = Get-NombreLimpio ([string]$cfg.Empresa)
    if (-not $empresaLimpia) { $empresaLimpia = 'Empresa' }
    $D.ClaveMarca    = 'HKLM:\SOFTWARE\' + $empresaLimpia + '\Paquetes\' + (Get-NombreLimpio $D.NombreApp)

    if (Test-Path -LiteralPath $D.RutaPaquete) {
        throw "Ya existe el paquete '$($D.RutaPaquete)'. Bórralo o cambia la versión."
    }
    if ($D.CopiarCarpetaCompleta) {
        $raizNorm = $script:Raiz.TrimEnd('\') + '\'
        if ($baseOrigen.StartsWith($raizNorm, [System.StringComparison]::OrdinalIgnoreCase) -or $raizNorm.StartsWith($baseOrigen, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'El instalador no puede estar dentro de la carpeta de la Fábrica (ni la Fábrica dentro de la carpeta del instalador). Muévelo, por ejemplo, a D:\Fuentes\<App>.'
        }
        $tamano = (Get-ChildItem -LiteralPath $carpetaOrigen -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $tamano) { $tamano = 0 }
        Write-Registro -Mensaje ('Se copiará la carpeta completa del instalador ({0:N0} MB).' -f ($tamano / 1MB))
    }
    if ($D.CopiarCarpetaCompleta) { Test-LongitudRutaPaquete -RutaPaquete $D.RutaPaquete -CarpetaOrigen $carpetaOrigen }
    else { Test-LongitudRutaPaquete -RutaPaquete $D.RutaPaquete }
    Write-Registro -Mensaje "Paquete: $($D.NombrePaquete)  (tipo $($D.Tipo))"

    try {
        # 3. Plantilla oficial PSADT
        Write-Registro -Mensaje 'Creando estructura desde la plantilla oficial de PSADT...'
        New-ADTTemplate -Destination $script:CarpetaPaquetes -Name $D.NombrePaquete | Out-Null
        $rutaScript = Join-Path $D.RutaPaquete 'Invoke-AppDeployToolkit.ps1'
        if (-not (Test-Path -LiteralPath $rutaScript)) { throw 'New-ADTTemplate no generó Invoke-AppDeployToolkit.ps1.' }

        # 4. Plantilla corporativa (logo, configuración, textos)
        Copy-RecursosCorporativos -RutaPaquete $D.RutaPaquete
        if ($D.LogoApp -and $D.LogoAppEnVentanas -and (Test-Path -LiteralPath $D.LogoApp)) {
            Set-LogoVentanas -RutaPaquete $D.RutaPaquete -RutaImagen $D.LogoApp
            Write-Registro -Mensaje 'Plantilla corporativa aplicada con el logo de la app en las ventanas.'
        }
        else {
            Write-Registro -Mensaje 'Plantilla corporativa aplicada (logo de la empresa en las ventanas).'
        }
        Set-TextosVentana -RutaPaquete $D.RutaPaquete -Mensaje $D.MensajeBienvenida -BotonInstalar $D.TextoBotonInstalar -BotonAplazar $D.TextoBotonAplazar

        # 5. Copiar instalador
        $carpetaFiles = Join-Path $D.RutaPaquete 'Files'
        Get-ChildItem -LiteralPath $carpetaFiles -Filter '*.txt' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'Add * Here.txt' } | Remove-Item -Force -ErrorAction SilentlyContinue
        if ($D.CopiarCarpetaCompleta) {
            Copy-Item -Path (Join-Path $carpetaOrigen '*') -Destination $carpetaFiles -Recurse -Force
            Write-Registro -Mensaje "Copiada la carpeta del instalador: $carpetaOrigen"
        }
        else {
            Copy-Item -LiteralPath $D.Instalador -Destination $carpetaFiles -Force
            Write-Registro -Mensaje "Copiado el instalador: $($D.RutaRelativaInstalador)"
        }

        # 6. Script de despliegue
        $valoresSesion = [ordered]@{
            AppVendor       = $D.Fabricante
            AppName         = $D.NombreApp
            AppVersion      = $D.Version
            AppArch         = $D.Arquitectura
            AppLang         = 'ES'
            AppRevision     = '01'
            AppScriptDate   = (Get-Date -Format 'yyyy-MM-dd')
            AppScriptAuthor = $env:USERNAME
        }
        $literalProcesos = Get-LiteralProcesos -Procesos $D.ProcesosACerrar -Descripcion $D.NombreApp
        $D.ProcesosEnSesion = $false
        if ($literalProcesos -ne '@()' -and -not $D.SiemprePreguntar) {
            if (Test-ProcesosEnSesion -RutaScript $rutaScript) { $D.ProcesosEnSesion = $true }
            else { Write-Registro -Mensaje 'Esta versión de PSADT no tiene AppProcessesToClose: los procesos se definen dentro de las fases.' -Tipo Aviso }
        }
        $fases = New-ContenidoFases -D $D
        if ($D.ProcesosEnSesion) {
            Update-ScriptDespliegue -RutaScript $rutaScript -ValoresSesion $valoresSesion -Fases $fases -ProcesosSesionLiteral $literalProcesos
        }
        else {
            Update-ScriptDespliegue -RutaScript $rutaScript -ValoresSesion $valoresSesion -Fases $fases
        }
        Write-Registro -Mensaje 'Invoke-AppDeployToolkit.ps1 generado (instalar, desinstalar, reparar).' -Tipo OK
    }
    catch {
        Write-Registro -Mensaje 'Falló la creación; se elimina la carpeta incompleta del paquete.' -Tipo Aviso
        if (Test-Path -LiteralPath $D.RutaPaquete) {
            Get-ChildItem -LiteralPath $D.RutaPaquete -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.PSIsContainer) { $_.IsReadOnly = $false } }
            Remove-Item -LiteralPath $D.RutaPaquete -Recurse -Force -ErrorAction SilentlyContinue
        }
        throw
    }

    # 7. Salida: detección, ficha y logo de la app
    if (-not (Test-Path -LiteralPath $D.RutaSalida)) { New-Item -Path $D.RutaSalida -ItemType Directory -Force | Out-Null }
    New-ScriptDeteccion -D $D -Ruta (Join-Path $D.RutaSalida 'Deteccion.ps1')

    $codigoAplazamiento = ''
    $rutaConfig = Join-Path $D.RutaPaquete 'Config\config.psd1'
    if (Test-Path -LiteralPath $rutaConfig) {
        $m = [regex]::Match([System.IO.File]::ReadAllText($rutaConfig), 'DeferExitCode\s*=\s*(\d+)')
        if ($m.Success) { $codigoAplazamiento = $m.Groups[1].Value }
    }
    New-FichaIntune -D $D -Ruta (Join-Path $D.RutaSalida 'FichaIntune.txt') -CodigoAplazamiento $codigoAplazamiento
    [System.IO.File]::WriteAllText((Join-Path $D.RutaSalida 'Descripcion.txt'), [string]$D.Descripcion, $script:Utf8ConBom)
    try {
        $rutaLogoIntune = Join-Path $D.RutaSalida 'Logo_Intune.png'
        $imgApp = $null
        if ($D.LogoApp -and (Test-Path -LiteralPath $D.LogoApp)) { $imgApp = Get-ImagenSinBloqueo $D.LogoApp }
        try {
            New-IconoPng -Ruta $rutaLogoIntune -Ancho 512 -Alto 512 -ColorHex ([string]$cfg.ColorAcento) -Iniciales (Get-Iniciales $D.NombreApp) -ImagenOrigen $imgApp
        }
        finally { if ($null -ne $imgApp) { $imgApp.Dispose() } }
    }
    catch { Write-Registro -Mensaje "No se pudo generar el logo de la app: $($_.Exception.Message)" -Tipo Aviso }
    Write-Registro -Mensaje 'Generados Deteccion.ps1, FichaIntune.txt, Descripcion.txt y Logo_Intune.png.' -Tipo OK

    # 8. .intunewin
    $intunewin = ''
    if ($D.GenerarIntunewin) {
        $util = Join-Path $script:Herramientas 'IntuneWinAppUtil.exe'
        if (-not (Test-Path -LiteralPath $util)) {
            Write-Registro -Mensaje 'No se generó el .intunewin: falta Herramientas\IntuneWinAppUtil.exe.' -Tipo Aviso
        }
        else {
            Write-Registro -Mensaje 'Generando .intunewin (puede tardar según el tamaño)...'
            $argumentos = @('-c', ('"{0}"' -f $D.RutaPaquete), '-s', 'Invoke-AppDeployToolkit.exe', '-o', ('"{0}"' -f $D.RutaSalida), '-q')
            $proceso = Start-Process -FilePath $util -ArgumentList $argumentos -Wait -PassThru -WindowStyle Hidden
            if ($proceso.ExitCode -ne 0) { throw "IntuneWinAppUtil.exe terminó con código $($proceso.ExitCode)." }
            $generado  = Join-Path $D.RutaSalida 'Invoke-AppDeployToolkit.intunewin'
            $intunewin = Join-Path $D.RutaSalida "$($D.NombrePaquete).intunewin"
            if (Test-Path -LiteralPath $intunewin) { Remove-Item -LiteralPath $intunewin -Force }
            if (Test-Path -LiteralPath $generado) {
                Rename-Item -LiteralPath $generado -NewName "$($D.NombrePaquete).intunewin"
                Write-Registro -Mensaje "Generado: $intunewin" -Tipo OK
            }
            else {
                Write-Registro -Mensaje 'IntuneWinAppUtil terminó pero no se encontró el .intunewin esperado.' -Tipo Aviso
                $intunewin = ''
            }
        }
    }

    if ($D.IncluirFuente) {
        $destinoFuente = Join-Path $D.RutaSalida 'Fuente'
        Write-Registro -Mensaje 'Copiando la carpeta fuente del paquete al destino...'
        if (Test-Path -LiteralPath $destinoFuente) {
            Get-ChildItem -LiteralPath $destinoFuente -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.PSIsContainer) { $_.IsReadOnly = $false } }
            Remove-Item -LiteralPath $destinoFuente -Recurse -Force
        }
        Copy-Item -LiteralPath $D.RutaPaquete -Destination $destinoFuente -Recurse -Force
        Write-Registro -Mensaje "Fuente copiada en: $destinoFuente" -Tipo OK
    }

    Add-RegistroHistorial -D $D -Intunewin $intunewin

    Write-Registro -Mensaje "Paquete listo en: $($D.RutaSalida)" -Tipo OK
    Write-Registro -Mensaje "Carpeta fuente (para editar): $($D.RutaPaquete)"
    Write-Registro -Mensaje 'Siguiente paso: probar como SYSTEM (ver FichaIntune.txt) antes de subir a Intune.'
    return @{ RutaPaquete = $D.RutaPaquete; RutaSalida = $D.RutaSalida; Intunewin = $intunewin }
}
#endregion

#region ===================== Historial y plantilla (compartido por escritorio y web) =====================
function Get-PropiedadJson {
    # Lee una propiedad de un objeto de ConvertFrom-Json sin fallar con Set-StrictMode si no existe.
    param($Objeto, [string]$Nombre, $PorDefecto = $null)
    if ($null -eq $Objeto) { return $PorDefecto }
    if ($Objeto -is [System.Collections.IDictionary]) {
        if ($Objeto.Contains($Nombre) -and $null -ne $Objeto[$Nombre]) { return $Objeto[$Nombre] }
        return $PorDefecto
    }
    $p = $Objeto.PSObject.Properties[$Nombre]
    if ($null -eq $p -or $null -eq $p.Value) { return $PorDefecto }
    return $p.Value
}

function Add-RegistroHistorial {
    # Anota el paquete creado en Recursos\Historial.json (lo usa la página «Paquetes creados»).
    param([Parameter(Mandatory)][hashtable]$D, [string]$Intunewin)
    try {
        $entradas = New-Object System.Collections.Generic.List[object]
        $entradas.Add([ordered]@{
                Fecha         = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                NombrePaquete = [string]$D.NombrePaquete
                NombreApp     = [string]$D.NombreApp
                Version       = [string]$D.Version
                Fabricante    = [string]$D.Fabricante
                Tipo          = [string]$D.Tipo
                RutaPaquete   = [string]$D.RutaPaquete
                RutaSalida    = [string]$D.RutaSalida
                Intunewin     = [string]$Intunewin
                Usuario       = [string]$env:USERNAME
            })
        if (Test-Path -LiteralPath $script:RutaHistorial) {
            $previo = ConvertFrom-Json ([System.IO.File]::ReadAllText($script:RutaHistorial))
            foreach ($e in $previo) {
                if ($null -eq $e) { continue }
                if ([string](Get-PropiedadJson $e 'RutaPaquete' '') -eq [string]$D.RutaPaquete) { continue }
                if ($entradas.Count -ge 300) { break }
                $entradas.Add($e)
            }
        }
        $json = ConvertTo-Json -InputObject ([object[]]$entradas.ToArray()) -Depth 4
        [System.IO.File]::WriteAllText($script:RutaHistorial, $json, $script:Utf8ConBom)
    }
    catch { Write-Registro -Mensaje "No se pudo anotar el paquete en el historial: $($_.Exception.Message)" -Tipo Aviso }
}

function Get-HistorialPaquetes {
    # Paquetes del historial más las carpetas de Paquetes\ que no están en él (creadas con versiones
    # anteriores de la Fábrica). Del más reciente al más antiguo.
    $lista = New-Object System.Collections.Generic.List[object]
    $vistos = @{}
    $campos = @('Fecha', 'NombrePaquete', 'NombreApp', 'Version', 'Fabricante', 'Tipo', 'RutaPaquete', 'RutaSalida', 'Intunewin', 'Usuario')
    if (Test-Path -LiteralPath $script:RutaHistorial) {
        try {
            $datos = ConvertFrom-Json ([System.IO.File]::ReadAllText($script:RutaHistorial))
            foreach ($e in $datos) {
                if ($null -eq $e) { continue }
                $item = [ordered]@{}
                foreach ($campo in $campos) { $item[$campo] = [string](Get-PropiedadJson $e $campo '') }
                $lista.Add($item)
                if ($item.RutaPaquete) { $vistos[$item.RutaPaquete.ToLowerInvariant()] = $true }
            }
        }
        catch { Write-Registro -Mensaje "No se pudo leer Historial.json: $($_.Exception.Message)" -Tipo Aviso }
    }
    if (Test-Path -LiteralPath $script:CarpetaPaquetes) {
        foreach ($dir in @(Get-ChildItem -LiteralPath $script:CarpetaPaquetes -Directory -ErrorAction SilentlyContinue | Sort-Object -Property CreationTime -Descending)) {
            if ($vistos.ContainsKey($dir.FullName.ToLowerInvariant())) { continue }
            $salidaPosible = Join-Path $script:CarpetaSalida $dir.Name
            if (-not (Test-Path -LiteralPath $salidaPosible)) { $salidaPosible = '' }
            $item = [ordered]@{}
            foreach ($campo in $campos) { $item[$campo] = '' }
            $item.Fecha = $dir.CreationTime.ToString('yyyy-MM-dd HH:mm')
            $item.NombrePaquete = $dir.Name
            $item.RutaPaquete = $dir.FullName
            $item.RutaSalida = $salidaPosible
            $lista.Add($item)
        }
    }
    foreach ($item in $lista) {
        $item.ExisteFuente = [bool]($item.RutaPaquete -and (Test-Path -LiteralPath $item.RutaPaquete))
        $item.ExisteSalida = [bool]($item.RutaSalida -and (Test-Path -LiteralPath $item.RutaSalida))
    }
    return , $lista.ToArray()
}

function Save-PlantillaCorporativa {
    # Guarda la plantilla corporativa (botón «Guardar plantilla» del escritorio y de la web).
    # LogoPendiente / BannerPendiente: $null = no cambia, '' = volver a lo automático, ruta = imagen nueva.
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Valores,
        $LogoPendiente = $null,
        $BannerPendiente = $null
    )
    $cfg = Get-ConfigCorporativa
    $empresa = ([string]$Valores.Empresa).Trim()
    if ([string]::IsNullOrWhiteSpace($empresa)) { throw 'Escribe el nombre de la empresa.' }
    $hex = ([string]$Valores.ColorAcento).Trim()
    if (-not $hex.StartsWith('#')) { $hex = '#' + $hex }
    if ($hex -notmatch '^#[0-9A-Fa-f]{6}$') { throw 'El color debe tener el formato #RRGGBB (ej. #2C2A29).' }
    foreach ($par in @(@($LogoPendiente, 'logo'), @($BannerPendiente, 'banner'))) {
        if ($par[0] -and -not (Test-Path -LiteralPath ([string]$par[0]) -PathType Leaf)) { throw "No existe la imagen del $($par[1]): $($par[0])" }
    }

    $cfg.Empresa            = $empresa
    $cfg.ColorAcento        = $hex.ToUpper()
    $cfg.EstiloDialogo      = if ([string]$Valores.EstiloDialogo -eq 'Classic') { 'Classic' } else { 'Fluent' }
    $cfg.Idioma             = switch ([string]$Valores.Idioma) { 'es' { 'es' } 'en' { 'en' } default { 'auto' } }
    $cfg.RutaLogs           = ([string]$Valores.RutaLogs).Trim()
    $cfg.VecesAplazar       = [Math]::Max(0, [Math]::Min(10, [int]$Valores.VecesAplazar))
    $cfg.DiasLimite         = [Math]::Max(0, [Math]::Min(30, [int]$Valores.DiasLimite))
    $cfg.CuentaRegresivaMin = [Math]::Max(0, [Math]::Min(60, [int]$Valores.CuentaRegresivaMin))
    $cfg.MensajeBienvenida  = ([string]$Valores.MensajeBienvenida).Trim()
    $cfg.MensajeFinal       = ([string]$Valores.MensajeFinal).Trim()
    $cfg.TextoBotonInstalar = ([string]$Valores.TextoBotonInstalar).Trim()
    $cfg.TextoBotonAplazar  = ([string]$Valores.TextoBotonAplazar).Trim()

    if ($null -ne $LogoPendiente) {
        if ($LogoPendiente) {
            # Se lee primero en memoria por si el archivo elegido es el mismo logo guardado antes.
            $bytes = [System.IO.File]::ReadAllBytes([string]$LogoPendiente)
            $nombreLogo = 'LogoEmpresa' + [System.IO.Path]::GetExtension([string]$LogoPendiente).ToLowerInvariant()
            Get-ChildItem -LiteralPath $script:Recursos -Filter 'LogoEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            [System.IO.File]::WriteAllBytes((Join-Path $script:Recursos $nombreLogo), $bytes)
            $cfg.Logo = $nombreLogo
        }
        else {
            Get-ChildItem -LiteralPath $script:Recursos -Filter 'LogoEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            $cfg.Logo = ''
        }
    }

    if ($null -ne $BannerPendiente) {
        if ($BannerPendiente) {
            $bytesBanner = [System.IO.File]::ReadAllBytes([string]$BannerPendiente)
            $nombreBanner = 'BannerEmpresa' + [System.IO.Path]::GetExtension([string]$BannerPendiente).ToLowerInvariant()
            Get-ChildItem -LiteralPath $script:Recursos -Filter 'BannerEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            [System.IO.File]::WriteAllBytes((Join-Path $script:Recursos $nombreBanner), $bytesBanner)
            $cfg.Banner = $nombreBanner
        }
        else {
            Get-ChildItem -LiteralPath $script:Recursos -Filter 'BannerEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            $cfg.Banner = ''
        }
    }

    $sinImportarLogo = ($null -ne $LogoPendiente)
    $sinImportarBanner = ($null -ne $BannerPendiente)
    Save-ConfigCorporativa -Cfg $cfg
    Update-PlantillaCorporativa -Cfg $cfg -NoImportarLogo:$sinImportarLogo -NoImportarBanner:$sinImportarBanner
    Write-Registro -Mensaje 'Plantilla guardada: logo, banner y configuración actualizados.' -Tipo OK
    Write-Registro -Mensaje 'Los paquetes creados ANTES no cambian: vuelve a crearlos para que tomen la nueva plantilla.' -Tipo Aviso
    return $cfg
}

function Get-ProblemasDatos {
    # Validación previa a «Crear paquete» (la misma en escritorio y web). Devuelve @{ Errores; Avisos },
    # cada uno con elementos @{ Campo; Texto }.
    param([Parameter(Mandatory)][hashtable]$Datos, [string]$Tipo)
    $errores = @(); $avisos = @()
    if ([string]::IsNullOrWhiteSpace($Datos.Instalador)) { $errores += @{ Campo = 'Instalador'; Texto = 'Elige el instalador.' } }
    if ([string]::IsNullOrWhiteSpace($Datos.NombreApp)) { $errores += @{ Campo = 'NombreApp'; Texto = 'Escribe el nombre de la aplicación.' } }
    if ([string]::IsNullOrWhiteSpace($Datos.Version)) {
        if ($Tipo -eq 'PS1') { $avisos += @{ Campo = 'Version'; Texto = 'Sin versión: se usará 1.0.0 para la marca de detección.' } }
        else { $avisos += @{ Campo = 'Version'; Texto = 'Sin versión: la detección buscará la app solo por su nombre.' } }
    }
    if ($Tipo -eq 'EXE') {
        if ([string]::IsNullOrWhiteSpace($Datos.DesinstaladorExe)) { $errores += @{ Campo = 'DesinstaladorExe'; Texto = 'El EXE necesita la ruta del desinstalador en el equipo.' } }
        if ([string]::IsNullOrWhiteSpace($Datos.ArgumentosInstalacion)) { $avisos += @{ Campo = 'ArgumentosInstalacion'; Texto = 'Sin argumentos silenciosos: si el instalador muestra ventanas, Intune quedará esperando.' } }
    }
    if ($Tipo -eq 'PS1' -and -not [string]::IsNullOrWhiteSpace($Datos.DesinstaladorExe) -and -not [string]::IsNullOrWhiteSpace($Datos.Instalador)) {
        $base = (Split-Path -Path $Datos.Instalador -Parent).TrimEnd('\') + '\'
        if (-not ([string]$Datos.DesinstaladorExe).StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
            $errores += @{ Campo = 'DesinstaladorExe'; Texto = 'El script de desinstalación debe estar en la carpeta del script de instalación.' }
        }
    }
    if ([string]::IsNullOrWhiteSpace($Datos.Descripcion)) { $avisos += @{ Campo = 'Descripcion'; Texto = 'Sin descripción: se escribirá una automática.' } }
    return @{ Errores = $errores; Avisos = $avisos }
}

function ConvertTo-DatosPaquete {
    # Convierte lo que envía el navegador (JSON) en el mismo hashtable que arma la interfaz de escritorio.
    param($Objeto)
    $texto = { param([string]$n) ([string](Get-PropiedadJson $Objeto $n '')).Trim() }
    $logico = { param([string]$n) [bool](Get-PropiedadJson $Objeto $n $false) }
    $veces = [Math]::Max(0, [Math]::Min(10, [int](Get-PropiedadJson $Objeto 'VecesAplazar' 0)))
    $arq = & $texto 'Arquitectura'
    if ($arq -ne 'x86') { $arq = 'x64' }
    return @{
        Instalador               = & $texto 'Instalador'
        Fabricante               = & $texto 'Fabricante'
        NombreApp                = & $texto 'NombreApp'
        Version                  = & $texto 'Version'
        Descripcion              = & $texto 'Descripcion'
        Arquitectura             = $arq
        ArgumentosInstalacion    = & $texto 'ArgumentosInstalacion'
        DesinstaladorExe         = & $texto 'DesinstaladorExe'
        ArgumentosDesinstalacion = & $texto 'ArgumentosDesinstalacion'
        NombreDeteccion          = & $texto 'NombreDeteccion'
        ProcesosACerrar          = & $texto 'ProcesosACerrar'
        SiemprePreguntar         = & $logico 'SiemprePreguntar'
        PermitirAplazar          = ($veces -gt 0)
        VecesAplazar             = [Math]::Max(1, $veces)
        DiasLimite               = [Math]::Max(0, [Math]::Min(30, [int](Get-PropiedadJson $Objeto 'DiasLimite' 0)))
        CuentaRegresivaMin       = [Math]::Max(0, [Math]::Min(60, [int](Get-PropiedadJson $Objeto 'CuentaRegresivaMin' 0)))
        MensajeBienvenida        = & $texto 'MensajeBienvenida'
        TextoBotonInstalar       = & $texto 'TextoBotonInstalar'
        TextoBotonAplazar        = & $texto 'TextoBotonAplazar'
        MensajeFinal             = & $texto 'MensajeFinal'
        LogoApp                  = & $texto 'LogoApp'
        LogoAppEnVentanas        = & $logico 'LogoAppEnVentanas'
        CopiarCarpetaCompleta    = & $logico 'CopiarCarpetaCompleta'
        GenerarIntunewin         = & $logico 'GenerarIntunewin'
        CarpetaDestino           = & $texto 'CarpetaDestino'
        IncluirFuente            = & $logico 'IncluirFuente'
    }
}
#endregion

#region ===================== Interfaz gráfica =====================
$script:XamlInterfaz = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Fábrica de Paquetes PSADT para Intune · v4.0" Width="1180" Height="820" MinWidth="960" MinHeight="600"
        WindowStartupLocation="CenterScreen" Background="#F4F6F9" FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True">
  <Window.Resources>
    <SolidColorBrush x:Key="Acento" Color="#2C2A29"/>
    <SolidColorBrush x:Key="AcentoTexto" Color="White"/>
    <SolidColorBrush x:Key="AcentoSuave" Color="#EDEDED"/>

    <Style x:Key="Tarjeta" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="BorderBrush" Value="#E4E7EC"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="20,16"/>
      <Setter Property="Margin" Value="0,0,0,14"/>
    </Style>
    <Style x:Key="TituloTarjeta" TargetType="TextBlock">
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#101828"/>
      <Setter Property="Margin" Value="0,0,0,2"/>
    </Style>
    <Style x:Key="SubtituloTarjeta" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#667085"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="Etiqueta" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#344054"/>
      <Setter Property="Margin" Value="0,10,0,4"/>
    </Style>
    <Style x:Key="Ayuda" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="Foreground" Value="#667085"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,3,0,0"/>
    </Style>
    <Style x:Key="TextoError" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="Foreground" Value="#D92D20"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
      <Setter Property="Visibility" Value="Collapsed"/>
    </Style>
    <Style x:Key="Paso" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="Foreground" Value="#98A2B3"/>
      <Setter Property="Margin" Value="6,3,0,3"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="Foreground" Value="#101828"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="b" Background="White" BorderBrush="#D0D5DD" BorderThickness="1" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource Acento}"/>
              </Trigger>
              <Trigger Property="Tag" Value="Aviso">
                <Setter TargetName="b" Property="BorderBrush" Value="#F79009"/>
                <Setter TargetName="b" Property="BorderThickness" Value="1.5"/>
              </Trigger>
              <Trigger Property="Tag" Value="Error">
                <Setter TargetName="b" Property="BorderBrush" Value="#D92D20"/>
                <Setter TargetName="b" Property="BorderThickness" Value="1.5"/>
                <Setter TargetName="b" Property="Background" Value="#FFFBFA"/>
              </Trigger>
              <Trigger Property="IsReadOnly" Value="True">
                <Setter TargetName="b" Property="Background" Value="#F2F4F7"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="b" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BotonPrimario" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource Acento}"/>
      <Setter Property="Foreground" Value="{DynamicResource AcentoTexto}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Padding" Value="18,10"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.88"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.72"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BotonSecundario" TargetType="Button">
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#344054"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="#D0D5DD" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#F2F4F7"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Background" Value="#E4E7EC"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BotonMenu" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#344054"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Padding" Value="14,11"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.85"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,10,0,0"/>
      <Setter Property="Foreground" Value="#344054"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="Height" Value="32"/>
    </Style>
    <Style TargetType="Slider">
      <Setter Property="IsSnapToTickEnabled" Value="True"/>
      <Setter Property="TickFrequency" Value="1"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="{DynamicResource Acento}"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="210"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <!-- Encabezado -->
    <Border Grid.ColumnSpan="2" Background="{DynamicResource Acento}" Padding="22,14">
      <Grid>
        <StackPanel Orientation="Horizontal">
          <Border Width="40" Height="40" CornerRadius="8" Background="White" Margin="0,0,14,0">
            <Image x:Name="imgEncabezado" Margin="5" Stretch="Uniform"/>
          </Border>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="Fábrica de Paquetes" FontSize="20" FontWeight="SemiBold" Foreground="{DynamicResource AcentoTexto}"/>
            <TextBlock x:Name="txtSubtituloApp" Text="PSAppDeployToolkit → Intune" FontSize="12" Opacity="0.85" Foreground="{DynamicResource AcentoTexto}"/>
          </StackPanel>
        </StackPanel>
        <Border HorizontalAlignment="Right" VerticalAlignment="Center" Background="#33FFFFFF" CornerRadius="12" Padding="12,4">
          <TextBlock x:Name="txtVersionPsadt" Text="PSADT: cargando..." Foreground="{DynamicResource AcentoTexto}" FontSize="12"/>
        </Border>
      </Grid>
    </Border>

    <!-- Menú lateral -->
    <Border Grid.Row="1" Grid.RowSpan="2" Background="White" BorderBrush="#E4E7EC" BorderThickness="0,0,1,0" Padding="12,16">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
      <StackPanel>
        <Button x:Name="navCrear" Style="{StaticResource BotonMenu}" Content="📦   Crear paquete"/>
        <Button x:Name="navPaquetes" Style="{StaticResource BotonMenu}" Content="🗂   Paquetes creados"/>
        <Button x:Name="navPlantilla" Style="{StaticResource BotonMenu}" Content="🎨   Plantilla corporativa"/>
        <Button x:Name="navAyuda" Style="{StaticResource BotonMenu}" Content="❓   Cómo funciona"/>
        <Separator Margin="0,14"/>
        <TextBlock Text="TU AVANCE" FontSize="11" FontWeight="SemiBold" Foreground="#98A2B3" Margin="6,0,0,6"/>
        <TextBlock x:Name="pasoInstalador" Style="{StaticResource Paso}" Text="○  1. Instalador"/>
        <TextBlock x:Name="pasoInfo" Style="{StaticResource Paso}" Text="○  2. Información de la app"/>
        <TextBlock x:Name="pasoInstalacion" Style="{StaticResource Paso}" Text="○  3. Cómo se instala"/>
        <TextBlock x:Name="pasoVentanas" Style="{StaticResource Paso}" Text="○  4. Ventanas del usuario"/>
        <Separator Margin="0,14"/>
        <Button x:Name="btnModoWeb" Style="{StaticResource BotonSecundario}" Content="🌐  Abrir versión web" HorizontalAlignment="Stretch"/>
        <TextBlock Style="{StaticResource Ayuda}" Margin="6,12,0,0" Text="Atajos:&#x0a;Ctrl+O  elegir instalador&#x0a;Ctrl+Enter  crear paquete&#x0a;F5  ver ventanas reales"/>
      </StackPanel>
      </ScrollViewer>
    </Border>

    <!-- ===================== Página: Crear paquete ===================== -->
    <Grid x:Name="pagCrear" Grid.Row="1" Grid.Column="1" Margin="18,16,18,8">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="400"/>
      </Grid.ColumnDefinitions>

      <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,0,14,0">
        <StackPanel>

          <!-- 1. Instalador -->
          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource TituloTarjeta}" Text="1 · Instalador"/>
              <TextBlock Style="{StaticResource SubtituloTarjeta}" Text="Arrastra aquí el instalador o haz clic para buscarlo. Se leen solos el nombre, el fabricante y la versión."/>
              <Border x:Name="zonaArrastre" AllowDrop="True" Cursor="Hand" Background="{DynamicResource AcentoSuave}" CornerRadius="10" Padding="16" MinHeight="96">
                <Grid>
                  <Rectangle RadiusX="10" RadiusY="10" Stroke="{DynamicResource Acento}" StrokeThickness="1.5" StrokeDashArray="5 4" Margin="-16"/>
                  <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
                    <TextBlock x:Name="txtZonaIcono" Text="⬇" FontSize="26" HorizontalAlignment="Center" Foreground="{DynamicResource Acento}"/>
                    <TextBlock x:Name="txtZonaTitulo" Text="Arrastra aquí el .msi, .exe o .ps1" FontSize="14" FontWeight="SemiBold" HorizontalAlignment="Center" Foreground="#101828"/>
                    <TextBlock x:Name="txtZonaDetalle" Text="o haz clic para buscarlo" FontSize="12" HorizontalAlignment="Center" Foreground="#667085" TextTrimming="CharacterEllipsis"/>
                  </StackPanel>
                </Grid>
              </Border>
              <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                <Border x:Name="bdTipo" Background="#F2F4F7" CornerRadius="10" Padding="10,3" Visibility="Collapsed">
                  <TextBlock x:Name="txtTipo" FontWeight="SemiBold" FontSize="12" Foreground="#344054"/>
                </Border>
                <TextBlock x:Name="txtTipoDetalle" Margin="10,0,0,0" VerticalAlignment="Center" Style="{StaticResource Ayuda}"/>
              </StackPanel>
              <CheckBox x:Name="chkCarpeta" Content="Copiar toda la carpeta del instalador (si trae archivos adicionales)"/>
            </StackPanel>
          </Border>

          <!-- 2. Información de la app -->
          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource TituloTarjeta}" Text="2 · Información de la app"/>
              <TextBlock Style="{StaticResource SubtituloTarjeta}" Text="Así aparece la aplicación en Intune y en Company Portal."/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="2*"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Nombre de la aplicación"/>
                  <TextBox x:Name="txtNombre"/>
                  <TextBlock x:Name="errNombre" Style="{StaticResource TextoError}"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Versión"/>
                  <TextBox x:Name="txtVersion"/>
                  <TextBlock x:Name="errVersion" Style="{StaticResource TextoError}"/>
                </StackPanel>
              </Grid>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="2*"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Fabricante (publicador)"/>
                  <TextBox x:Name="txtFabricante"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Arquitectura"/>
                  <ComboBox x:Name="cmbArq" SelectedIndex="0">
                    <ComboBoxItem Content="x64"/>
                    <ComboBoxItem Content="x86"/>
                  </ComboBox>
                </StackPanel>
              </Grid>
              <TextBlock Style="{StaticResource Etiqueta}" Text="Descripción para Company Portal"/>
              <TextBox x:Name="txtDescripcion" Height="64" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" VerticalContentAlignment="Top"/>
              <TextBlock Style="{StaticResource Ayuda}" Text="Qué hace la aplicación y para quién es. Intune la exige; si la dejas vacía se escribe una automática. Queda en FichaIntune.txt y en Descripcion.txt para copiar y pegar."/>

              <TextBlock Style="{StaticResource Etiqueta}" Text="Logo de la app"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Border Width="60" Height="60" CornerRadius="10" BorderBrush="#E4E7EC" BorderThickness="1" Background="#F9FAFB">
                  <Grid>
                    <TextBlock x:Name="txtLogoAppIniciales" Text="AP" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="#98A2B3"/>
                    <Image x:Name="imgLogoApp" Margin="6" Stretch="Uniform"/>
                  </Grid>
                </Border>
                <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
                  <TextBlock x:Name="txtLogoAppInfo" Style="{StaticResource Ayuda}" Margin="0,0,0,6" Text="Sin logo: se generan las iniciales con el color de la empresa."/>
                  <StackPanel Orientation="Horizontal">
                    <Button x:Name="btnLogoApp" Style="{StaticResource BotonSecundario}" Content="Elegir imagen..." Margin="0,0,8,0"/>
                    <Button x:Name="btnQuitarLogoApp" Style="{StaticResource BotonSecundario}" Content="Quitar"/>
                  </StackPanel>
                  <CheckBox x:Name="chkLogoVentanas" Content="Mostrar este logo en las ventanas (en vez del logo de la empresa)"/>
                </StackPanel>
              </Grid>
            </StackPanel>
          </Border>

          <!-- 3. Instalación -->
          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource TituloTarjeta}" Text="3 · Cómo se instala"/>
              <TextBlock x:Name="txtSubInstalacion" Style="{StaticResource SubtituloTarjeta}" Text="Primero elige un instalador."/>
              <StackPanel x:Name="filaArgs" Visibility="Collapsed">
                <TextBlock x:Name="lblArgs" Style="{StaticResource Etiqueta}" Text="Argumentos de instalación"/>
                <TextBox x:Name="txtArgs"/>
                <TextBlock x:Name="errArgs" Style="{StaticResource TextoError}" Foreground="#B54708"/>
                <TextBlock x:Name="lblAyudaArgs" Style="{StaticResource Ayuda}"/>
              </StackPanel>
              <StackPanel x:Name="filaDes" Visibility="Collapsed">
                <TextBlock x:Name="lblDes" Style="{StaticResource Etiqueta}" Text="Desinstalador"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBox x:Name="txtDes"/>
                  <Button x:Name="btnDes" Grid.Column="1" Style="{StaticResource BotonSecundario}" Content="Buscar..." Margin="8,0,0,0"/>
                </Grid>
                <TextBlock x:Name="errDes" Style="{StaticResource TextoError}"/>
                <TextBlock x:Name="lblAyudaDes" Style="{StaticResource Ayuda}"/>
              </StackPanel>
              <StackPanel x:Name="filaArgsDes" Visibility="Collapsed">
                <TextBlock Style="{StaticResource Etiqueta}" Text="Argumentos de desinstalación"/>
                <TextBox x:Name="txtArgsDes"/>
              </StackPanel>
              <StackPanel x:Name="filaDet" Visibility="Collapsed">
                <TextBlock Style="{StaticResource Etiqueta}" Text="Nombre para la detección"/>
                <TextBox x:Name="txtDeteccion"/>
                <TextBlock Style="{StaticResource Ayuda}" Text="Como aparece en Configuración → Aplicaciones (Programas y características)."/>
              </StackPanel>
            </StackPanel>
          </Border>

          <!-- 4. Ventanas -->
          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource TituloTarjeta}" Text="4 · Ventanas que ve el usuario"/>
              <TextBlock Style="{StaticResource SubtituloTarjeta}" Text="Mira cómo cambia la vista previa de la derecha mientras ajustas."/>
              <TextBlock Style="{StaticResource Etiqueta}" Text="Aplicaciones que se deben cerrar"/>
              <TextBox x:Name="txtProcesos"/>
              <TextBlock Style="{StaticResource Ayuda}" Text="Separadas por coma, sin .exe. Puedes darles nombre visible: excel=Microsoft Excel, AppCaja=Aplicativo Caja"/>
              <CheckBox x:Name="chkSiempre" Content="Mostrar las ventanas aunque la app esté cerrada (si no, se instala en silencio)"/>

              <Grid Margin="0,14,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="190"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="110"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Veces que puede aplazar" VerticalAlignment="Center" FontWeight="SemiBold" Foreground="#344054"/>
                <Slider x:Name="sldAplazar" Grid.Column="1" Minimum="0" Maximum="10" Value="3"/>
                <TextBlock x:Name="lblAplazar" Grid.Column="2" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="#101828" FontWeight="SemiBold"/>

                <TextBlock Grid.Row="1" Text="Fecha límite para aplazar" VerticalAlignment="Center" FontWeight="SemiBold" Foreground="#344054"/>
                <Slider x:Name="sldDias" Grid.Row="1" Grid.Column="1" Minimum="0" Maximum="30" Value="0"/>
                <TextBlock x:Name="lblDias" Grid.Row="1" Grid.Column="2" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="#101828" FontWeight="SemiBold"/>

                <TextBlock Grid.Row="2" Text="Cuenta regresiva" VerticalAlignment="Center" FontWeight="SemiBold" Foreground="#344054"/>
                <Slider x:Name="sldCuenta" Grid.Row="2" Grid.Column="1" Minimum="0" Maximum="60" Value="0"/>
                <TextBlock x:Name="lblCuenta" Grid.Row="2" Grid.Column="2" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="#101828" FontWeight="SemiBold"/>
              </Grid>
              <TextBlock Style="{StaticResource Ayuda}" Text="Cuenta regresiva: al llegar a cero se cierran las apps y se instala solo."/>

              <TextBlock Style="{StaticResource Etiqueta}" Text="Texto de la ventana (opcional)"/>
              <TextBox x:Name="txtBienvenida" Height="56" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" VerticalContentAlignment="Top"/>
              <TextBlock Style="{StaticResource Ayuda}" Text="Si escribes algo, REEMPLAZA el texto estándar de PSADT. Vacío = texto estándar. Pulsa Enter para un salto de línea. Formato: [bold]negrita[/bold]  [italic]cursiva[/italic]  [accent]color[/accent]  [url]https://...[/url]"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Texto del botón Instalar"/>
                  <TextBox x:Name="txtBotonInstalar"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Texto del botón Aplazar"/>
                  <TextBox x:Name="txtBotonAplazar"/>
                </StackPanel>
              </Grid>
              <TextBlock Style="{StaticResource Ayuda}" Text="Vacío = texto estándar de PSADT («Instalar» / «Cerrar Aplicaciones e Instalar» y «Aplazar»). Usa textos cortos: el botón es angosto."/>
              <TextBlock Style="{StaticResource Etiqueta}" Text="Mensaje al terminar (opcional)"/>
              <TextBox x:Name="txtMensaje"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>

      <!-- Panel derecho: vista previa y acciones -->
      <DockPanel Grid.Column="1">
        <StackPanel DockPanel.Dock="Bottom">
          <Border Background="White" CornerRadius="8" BorderBrush="#E4E7EC" BorderThickness="1" Padding="12,10" Margin="0,0,0,10">
            <StackPanel>
              <TextBlock Text="📁  Dónde queda el paquete listo" FontWeight="SemiBold" Foreground="#344054"/>
              <TextBlock x:Name="txtDestino" FontSize="11.5" Foreground="#667085" TextTrimming="CharacterEllipsis" Margin="0,4,0,8"/>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="btnDestino" Style="{StaticResource BotonSecundario}" Content="Cambiar carpeta..." Margin="0,0,8,0"/>
                <Button x:Name="btnDestinoDefecto" Style="{StaticResource BotonSecundario}" Content="Usar la de la Fábrica"/>
              </StackPanel>
              <CheckBox x:Name="chkCopiarFuente" Content="Incluir también la carpeta fuente (para editarlo después)" Margin="0,8,0,0"/>
            </StackPanel>
          </Border>
          <CheckBox x:Name="chkIntunewin" Content="Generar también el archivo .intunewin" IsChecked="True" Margin="2,0,0,10"/>
          <Border x:Name="bdResumen" CornerRadius="8" Padding="12,8" Margin="0,0,0,10" Background="#F2F4F7">
            <TextBlock x:Name="txtResumen" TextWrapping="Wrap" FontSize="12" Foreground="#344054" Text="Arrastra un instalador para empezar."/>
          </Border>
          <Button x:Name="btnCrear" Style="{StaticResource BotonPrimario}" Content="Crear paquete" Margin="0,0,0,8"/>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="8"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="btnVista" Style="{StaticResource BotonSecundario}" Content="👁  Ver ventanas reales"/>
            <Button x:Name="btnAbrir" Grid.Column="2" Style="{StaticResource BotonSecundario}" Content="📂  Abrir salida" IsEnabled="False"/>
          </Grid>
        </StackPanel>

        <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="0,0,6,10">
          <TextBlock Text="Así lo verá el usuario" FontSize="15" FontWeight="SemiBold" Foreground="#101828"/>
          <TextBlock Style="{StaticResource Ayuda}" Margin="0,2,0,10" Text="Vista aproximada. Para ver las ventanas reales de PSADT usa «Ver ventanas reales» (no instala nada)."/>
          <Border Background="White" CornerRadius="8" BorderBrush="#D0D5DD" BorderThickness="1">
            <Border.Effect>
              <DropShadowEffect BlurRadius="18" ShadowDepth="3" Opacity="0.18"/>
            </Border.Effect>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Border Background="{DynamicResource Acento}" CornerRadius="8,0,0,8"/>
              <StackPanel Grid.Column="1" Margin="14,14,14,14">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid Width="38" Height="38">
                    <Border x:Name="mockIniciales" CornerRadius="8" Background="{DynamicResource Acento}">
                      <TextBlock x:Name="mockInicialesTexto" Text="AP" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource AcentoTexto}"/>
                    </Border>
                    <Image x:Name="mockLogo" Stretch="Uniform"/>
                  </Grid>
                  <StackPanel Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center">
                    <TextBlock x:Name="mockTitulo" Text="Nombre de la app 1.0" FontWeight="SemiBold" FontSize="15" Foreground="#101828" TextTrimming="CharacterEllipsis"/>
                    <TextBlock x:Name="mockSubtitulo" Text="Empresa · Instalación de la aplicación" FontSize="11" Foreground="#667085"/>
                  </StackPanel>
                </Grid>
                <Border BorderBrush="#E4E7EC" BorderThickness="1" CornerRadius="6" Margin="0,12,0,0" Padding="10">
                  <StackPanel>
                    <TextBlock x:Name="mockTexto" TextWrapping="Wrap" FontSize="11.5" Foreground="#344054"/>
                    <TextBlock x:Name="mockMensaje" TextWrapping="Wrap" FontSize="11.5" Foreground="#344054" Margin="0,8,0,0"/>
                    <Border x:Name="mockListaBorde" BorderBrush="#E4E7EC" BorderThickness="1" CornerRadius="6" Margin="0,10,0,0" Padding="8,4">
                      <StackPanel x:Name="mockLista"/>
                    </Border>
                  </StackPanel>
                </Border>
                <StackPanel x:Name="mockFilas" Margin="0,6,0,0"/>
                <Grid Margin="0,12,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Border x:Name="mockBtnInstalar" BorderBrush="#D0D5DD" BorderThickness="1" CornerRadius="4" Padding="6">
                    <TextBlock x:Name="mockBtnInstalarTexto" Text="Cerrar apps e instalar" HorizontalAlignment="Center" FontSize="11.5" Foreground="#344054"/>
                  </Border>
                  <Border x:Name="mockBtnAplazar" Grid.Column="2" Background="{DynamicResource Acento}" CornerRadius="4" Padding="6">
                    <TextBlock x:Name="mockBtnAplazarTexto" Text="Aplazar" HorizontalAlignment="Center" FontSize="11.5" Foreground="{DynamicResource AcentoTexto}"/>
                  </Border>
                </Grid>
              </StackPanel>
            </Grid>
          </Border>
          <Border x:Name="bdAvisoSilencio" Background="#FFF6E5" CornerRadius="8" Padding="12,8" Margin="0,12,0,0">
            <TextBlock x:Name="txtAvisoSilencio" Style="{StaticResource Ayuda}" Margin="0" Foreground="#93370D"/>
          </Border>
        </StackPanel>
        </ScrollViewer>
      </DockPanel>
    </Grid>

    <!-- ===================== Página: Paquetes creados ===================== -->
    <Grid x:Name="pagPaquetes" Grid.Row="1" Grid.Column="1" Margin="18,16,18,8" Visibility="Collapsed">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Grid Margin="0,0,0,12">
        <StackPanel>
          <TextBlock Style="{StaticResource TituloTarjeta}" FontSize="18" Text="Paquetes creados"/>
          <TextBlock x:Name="txtPaquetesResumen" Style="{StaticResource Ayuda}" Text="Los paquetes que se han creado con esta Fábrica, del más reciente al más antiguo."/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
          <TextBox x:Name="txtBuscarPaquete" Width="240" Margin="0,0,8,0" ToolTip="Filtrar por nombre, versión o fabricante"/>
          <Button x:Name="btnRefrescarPaquetes" Style="{StaticResource BotonSecundario}" Content="⟳  Actualizar"/>
        </StackPanel>
      </Grid>
      <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
        <StackPanel x:Name="pnlPaquetes"/>
      </ScrollViewer>
    </Grid>

    <!-- ===================== Página: Plantilla corporativa ===================== -->
    <Grid x:Name="pagPlantilla" Grid.Row="1" Grid.Column="1" Margin="18,16,18,8" Visibility="Collapsed">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="400"/>
      </Grid.ColumnDefinitions>
      <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="0,0,14,0">
        <StackPanel>
          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource TituloTarjeta}" Text="Identidad de la empresa"/>
              <TextBlock Style="{StaticResource SubtituloTarjeta}" Text="Se aplica a todos los paquetes que crees DESPUÉS de guardar."/>
              <TextBlock Style="{StaticResource Etiqueta}" Text="Nombre de la empresa"/>
              <TextBox x:Name="txtEmpresa"/>
              <TextBlock Style="{StaticResource Ayuda}" Text="Sale en el subtítulo de las ventanas y en el banner."/>

              <TextBlock Style="{StaticResource Etiqueta}" Text="Color corporativo"/>
              <StackPanel Orientation="Horizontal">
                <Border x:Name="bdColor" Width="34" Height="32" CornerRadius="6" BorderBrush="#D0D5DD" BorderThickness="1" Margin="0,0,8,0"/>
                <TextBox x:Name="txtColor" Width="100" Margin="0,0,8,0"/>
                <Button x:Name="btnColor" Style="{StaticResource BotonSecundario}" Content="Elegir..." Margin="0,0,14,0"/>
                <WrapPanel x:Name="pnlPaleta" VerticalAlignment="Center"/>
              </StackPanel>

              <TextBlock Style="{StaticResource Etiqueta}" Text="Logo de la empresa"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Border Width="60" Height="60" CornerRadius="10" BorderBrush="#E4E7EC" BorderThickness="1" Background="#F9FAFB">
                  <Image x:Name="imgLogoEmp" Margin="6" Stretch="Uniform"/>
                </Border>
                <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
                  <TextBlock x:Name="txtLogoEmpInfo" Style="{StaticResource Ayuda}" Margin="0,0,0,6"/>
                  <StackPanel Orientation="Horizontal">
                    <Button x:Name="btnLogo" Style="{StaticResource BotonSecundario}" Content="Elegir imagen..." Margin="0,0,8,0"/>
                    <Button x:Name="btnQuitarLogo" Style="{StaticResource BotonSecundario}" Content="Usar iniciales"/>
                  </StackPanel>
                </StackPanel>
              </Grid>
              <TextBlock Style="{StaticResource Ayuda}" Text="Ideal: PNG cuadrado con fondo transparente, 256 px o más."/>

              <TextBlock Style="{StaticResource Etiqueta}" Text="Banner (solo se ve con el estilo Classic)"/>
              <Border CornerRadius="6" BorderBrush="#E4E7EC" BorderThickness="1" Background="#F9FAFB" HorizontalAlignment="Left" Padding="4">
                <Image x:Name="imgBannerEmp" Width="360" Height="50" Stretch="Uniform"/>
              </Border>
              <TextBlock x:Name="txtBannerEmpInfo" Style="{StaticResource Ayuda}" Margin="0,6,0,6"/>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="btnBanner" Style="{StaticResource BotonSecundario}" Content="Subir banner..." Margin="0,0,8,0"/>
                <Button x:Name="btnQuitarBanner" Style="{StaticResource BotonSecundario}" Content="Generar automático"/>
              </StackPanel>
              <TextBlock Style="{StaticResource Ayuda}" Text="Recomendado: PNG de 900 × 125 px (proporción 7,2 a 1). Con otra proporción se centra sin deformarse."/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource TituloTarjeta}" Text="Ventanas"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Estilo"/>
                  <ComboBox x:Name="cmbEstilo">
                    <ComboBoxItem Content="Fluent"/>
                    <ComboBoxItem Content="Classic"/>
                  </ComboBox>
                  <TextBlock Style="{StaticResource Ayuda}" Text="Fluent = moderno (recomendado)."/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Idioma"/>
                  <ComboBox x:Name="cmbIdioma">
                    <ComboBoxItem Content="Español"/>
                    <ComboBoxItem Content="Inglés"/>
                    <ComboBoxItem Content="Automático (idioma del equipo)"/>
                  </ComboBox>
                </StackPanel>
              </Grid>
              <TextBlock Style="{StaticResource Etiqueta}" Text="Ruta de logs en el equipo"/>
              <TextBox x:Name="txtLogs"/>
              <TextBlock Style="{StaticResource Ayuda}" Text="Por defecto $envWinDir\Logs\Software  (C:\Windows\Logs\Software)."/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource TituloTarjeta}" Text="Valores por defecto de cada paquete nuevo"/>
              <Grid Margin="0,8,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="190"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="110"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                  <RowDefinition Height="36"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Veces que puede aplazar" VerticalAlignment="Center" FontWeight="SemiBold" Foreground="#344054"/>
                <Slider x:Name="sldAplazarDef" Grid.Column="1" Minimum="0" Maximum="10"/>
                <TextBlock x:Name="lblAplazarDef" Grid.Column="2" VerticalAlignment="Center" Margin="12,0,0,0" FontWeight="SemiBold"/>
                <TextBlock Grid.Row="1" Text="Fecha límite para aplazar" VerticalAlignment="Center" FontWeight="SemiBold" Foreground="#344054"/>
                <Slider x:Name="sldDiasDef" Grid.Row="1" Grid.Column="1" Minimum="0" Maximum="30"/>
                <TextBlock x:Name="lblDiasDef" Grid.Row="1" Grid.Column="2" VerticalAlignment="Center" Margin="12,0,0,0" FontWeight="SemiBold"/>
                <TextBlock Grid.Row="2" Text="Cuenta regresiva" VerticalAlignment="Center" FontWeight="SemiBold" Foreground="#344054"/>
                <Slider x:Name="sldCuentaDef" Grid.Row="2" Grid.Column="1" Minimum="0" Maximum="60"/>
                <TextBlock x:Name="lblCuentaDef" Grid.Row="2" Grid.Column="2" VerticalAlignment="Center" Margin="12,0,0,0" FontWeight="SemiBold"/>
              </Grid>
              <TextBlock Style="{StaticResource Etiqueta}" Text="Texto de la ventana por defecto (reemplaza el estándar de PSADT)"/>
              <TextBox x:Name="txtBienvenidaDef" Height="56" TextWrapping="Wrap" AcceptsReturn="True" VerticalContentAlignment="Top"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Botón Instalar por defecto"/>
                  <TextBox x:Name="txtBotonInstalarDef"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Botón Aplazar por defecto"/>
                  <TextBox x:Name="txtBotonAplazarDef"/>
                </StackPanel>
              </Grid>
              <TextBlock Style="{StaticResource Etiqueta}" Text="Mensaje final por defecto"/>
              <TextBox x:Name="txtMensajeDef"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>

      <DockPanel Grid.Column="1">
        <StackPanel DockPanel.Dock="Bottom">
          <Button x:Name="btnGuardar" Style="{StaticResource BotonPrimario}" Content="Guardar plantilla" Margin="0,0,0,8"/>
          <Button x:Name="btnAbrirRecursos" Style="{StaticResource BotonSecundario}" Content="📂  Abrir carpeta Recursos"/>
        </StackPanel>
        <StackPanel>
          <TextBlock Text="Así quedan los recursos" FontSize="15" FontWeight="SemiBold" Foreground="#101828"/>
          <TextBlock Style="{StaticResource Ayuda}" Margin="0,2,0,12" Text="Se regeneran al pulsar «Guardar plantilla»."/>
          <Border Style="{StaticResource Tarjeta}">
            <StackPanel>
              <TextBlock Style="{StaticResource Etiqueta}" Margin="0,0,0,8" Text="Logo de las ventanas (Fluent)"/>
              <Border Width="110" Height="110" HorizontalAlignment="Left" CornerRadius="10" Background="#F9FAFB" BorderBrush="#E4E7EC" BorderThickness="1">
                <Image x:Name="imgIconoPrev" Margin="8" Stretch="Uniform"/>
              </Border>
              <TextBlock Style="{StaticResource Etiqueta}" Margin="0,14,0,8" Text="Banner (estilo Classic)"/>
              <Image x:Name="imgBannerPrev" Stretch="Uniform" HorizontalAlignment="Left" MaxHeight="60"/>
            </StackPanel>
          </Border>
          <Border x:Name="bdPendiente" Background="#FFF6E5" CornerRadius="8" Padding="12,8" Visibility="Collapsed">
            <TextBlock Style="{StaticResource Ayuda}" Margin="0" Foreground="#93370D" Text="Tienes cambios sin guardar. Pulsa «Guardar plantilla»."/>
          </Border>
        </StackPanel>
      </DockPanel>
    </Grid>

    <!-- ===================== Página: Cómo funciona ===================== -->
    <ScrollViewer x:Name="pagAyuda" Grid.Row="1" Grid.Column="1" Margin="18,16,18,8" Visibility="Collapsed" VerticalScrollBarVisibility="Auto">
      <StackPanel MaxWidth="860" HorizontalAlignment="Left">
        <Border Style="{StaticResource Tarjeta}">
          <StackPanel>
            <TextBlock Style="{StaticResource TituloTarjeta}" Text="Crear un paquete en 4 pasos"/>
            <TextBlock Style="{StaticResource Ayuda}" FontSize="13" Text="1. Arrastra el instalador (.msi, .exe o .ps1) a la zona punteada.&#x0a;2. Revisa nombre, versión, fabricante y escribe la descripción para Company Portal.&#x0a;3. Ajusta las ventanas (apps a cerrar, aplazamientos, cuenta regresiva) y mira la vista previa.&#x0a;4. Pulsa «Crear paquete». En Salida\&lt;paquete&gt; quedan el .intunewin, FichaIntune.txt, Deteccion.ps1, Descripcion.txt y Logo_Intune.png."/>
          </StackPanel>
        </Border>
        <Border Style="{StaticResource Tarjeta}">
          <StackPanel>
            <TextBlock Style="{StaticResource TituloTarjeta}" Text="¿Por qué a veces no salen las ventanas?"/>
            <TextBlock Style="{StaticResource Ayuda}" FontSize="13" Text="PSADT decide solo (modo Auto):&#x0a;• Si hay apps a cerrar y están ABIERTAS → muestra la ventana completa.&#x0a;• Si hay apps a cerrar y están CERRADAS → instala en silencio (no molesta al usuario).&#x0a;• Si no hay nadie conectado → siempre en silencio.&#x0a;Para que salgan siempre, marca «Mostrar las ventanas aunque la app esté cerrada»."/>
          </StackPanel>
        </Border>
        <Border Style="{StaticResource Tarjeta}">
          <StackPanel>
            <TextBlock Style="{StaticResource TituloTarjeta}" Text="Tipos de instalador"/>
            <TextBlock Style="{StaticResource Ayuda}" FontSize="13" Text="• MSI: todo es automático. Se detecta por su código de producto.&#x0a;• EXE: escribe los argumentos silenciosos (/S, /quiet...) y la ruta del desinstalador en el equipo.&#x0a;• PS1: se copia toda la carpeta del script; debe terminar con «exit 0». Se detecta por una marca en el registro que crea el paquete."/>
          </StackPanel>
        </Border>
        <Border Style="{StaticResource Tarjeta}">
          <StackPanel>
            <TextBlock Style="{StaticResource TituloTarjeta}" Text="Novedades de la versión 4.0"/>
            <TextBlock Style="{StaticResource Ayuda}" FontSize="13" Text="• Los campos obligatorios se marcan en rojo mientras escribes, y el menú de la izquierda muestra tu avance.&#x0a;• «Paquetes creados» lista lo que se ha hecho con la Fábrica y abre la salida o la carpeta fuente de cada uno.&#x0a;• «Abrir versión web» abre la misma Fábrica en el navegador (solo en este equipo). También con Crear-Paquete-Web.cmd.&#x0a;• Atajos: Ctrl+O elegir instalador, Ctrl+Enter crear, F5 ver ventanas reales."/>
          </StackPanel>
        </Border>
        <Border Style="{StaticResource Tarjeta}">
          <StackPanel>
            <TextBlock Style="{StaticResource TituloTarjeta}" Text="Antes de publicar en Intune"/>
            <TextBlock Style="{StaticResource Ayuda}" FontSize="13" Text="Prueba como SYSTEM con los comandos de FichaIntune.txt (psexec -s -i cmd). Instala, revisa que Deteccion.ps1 diga «Detectado», desinstala y vuelve a revisar. Luego publica en un grupo piloto."/>
          </StackPanel>
        </Border>
      </StackPanel>
    </ScrollViewer>

    <!-- Barra inferior: estado y registro -->
    <Border Grid.Row="2" Grid.Column="1" Background="White" BorderBrush="#E4E7EC" BorderThickness="0,1,0,0" Padding="18,8">
      <Expander x:Name="expRegistro">
        <Expander.Header>
          <TextBlock x:Name="txtEstado" Text="Preparando..." Foreground="#667085" TextTrimming="CharacterEllipsis"/>
        </Expander.Header>
        <TextBox x:Name="txtLog" Height="130" Margin="0,8,0,0" IsReadOnly="True" FontFamily="Consolas" FontSize="11.5" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
      </Expander>
    </Border>
  </Grid>
</Window>
'@

function Get-ImagenWpf {
    # Carga una imagen para WPF sin dejar el archivo bloqueado.
    param([string]$Ruta)
    if ([string]::IsNullOrWhiteSpace($Ruta) -or -not (Test-Path -LiteralPath $Ruta)) { return $null }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Ruta)
        $ms = New-Object System.IO.MemoryStream(, $bytes)
        $bi = New-Object System.Windows.Media.Imaging.BitmapImage
        $bi.BeginInit()
        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bi.StreamSource = $ms
        $bi.EndInit()
        $bi.Freeze()
        $ms.Dispose()
        return $bi
    }
    catch { return $null }
}

function Get-RutaLogoCorporativo {
    # Primer PNG cuadrado de Recursos\Assets (el logo que usan las ventanas).
    $carpeta = Join-Path $script:Recursos 'Assets'
    if (-not (Test-Path -LiteralPath $carpeta)) { return '' }
    foreach ($png in @(Get-ChildItem -LiteralPath $carpeta -Filter '*.png' -File)) {
        $img = Get-ImagenSinBloqueo $png.FullName
        if ($null -eq $img) { continue }
        $cuadrado = (($img.Width / [double]$img.Height) -lt 3)
        $img.Dispose()
        if ($cuadrado) { return $png.FullName }
    }
    return ''
}

function Get-RutaBannerCorporativo {
    $carpeta = Join-Path $script:Recursos 'Assets'
    if (-not (Test-Path -LiteralPath $carpeta)) { return '' }
    foreach ($png in @(Get-ChildItem -LiteralPath $carpeta -Filter '*.png' -File)) {
        $img = Get-ImagenSinBloqueo $png.FullName
        if ($null -eq $img) { continue }
        $banner = (($img.Width / [double]$img.Height) -ge 3)
        $img.Dispose()
        if ($banner) { return $png.FullName }
    }
    return ''
}

function Set-TextoFormateado {
    # Muestra en un TextBlock el texto con las etiquetas de PSADT: [bold] [italic] [accent] [url].
    param($Bloque, [string]$Texto)
    $Bloque.Inlines.Clear()
    if ([string]::IsNullOrEmpty($Texto)) { return }
    $negrita = 0; $cursiva = 0; $acento = 0
    $rx = [regex]'\[url=(?<u>[^\]]+)\](?<d>.+?)\[/url\]|\[url\](?<s>.+?)\[/url\]|\[(?<tag>/?(?:bold|italic|accent))\]'
    $agregar = {
        param([string]$t, [bool]$enlace)
        if (-not $t) { return }
        $run = New-Object System.Windows.Documents.Run($t)
        if ($negrita -gt 0) { $run.FontWeight = [System.Windows.FontWeights]::Bold }
        if ($cursiva -gt 0) { $run.FontStyle = [System.Windows.FontStyles]::Italic }
        if ($acento -gt 0 -or $enlace) { $run.Foreground = $Bloque.TryFindResource('Acento') }
        if ($enlace) { $run.TextDecorations = [System.Windows.TextDecorations]::Underline }
        $Bloque.Inlines.Add($run)
    }
    $pos = 0
    foreach ($m in $rx.Matches($Texto)) {
        if ($m.Index -gt $pos) { & $agregar $Texto.Substring($pos, $m.Index - $pos) $false }
        if ($m.Groups['tag'].Success) {
            switch ($m.Groups['tag'].Value) {
                'bold' { $negrita++ }
                '/bold' { $negrita-- }
                'italic' { $cursiva++ }
                '/italic' { $cursiva-- }
                'accent' { $acento++ }
                '/accent' { $acento-- }
            }
        }
        elseif ($m.Groups['d'].Success) { & $agregar $m.Groups['d'].Value $true }
        else { & $agregar $m.Groups['s'].Value $true }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $Texto.Length) { & $agregar $Texto.Substring($pos) $false }
}

function New-BrochaWpf {
    # Crea un relleno (Brush) de WPF a partir de un color #RRGGBB. Se usa BrushConverter con texto
    # para evitar problemas de conversión de tipos entre PowerShell y WPF.
    param([string]$Hex, [int]$Alfa = 255)
    $col = ConvertFrom-ColorHex $Hex
    $texto = '#{0:X2}{1:X2}{2:X2}{3:X2}' -f $Alfa, $col.R, $col.G, $col.B
    $conversor = New-Object System.Windows.Media.BrushConverter
    $brocha = [System.Windows.Media.Brush]$conversor.ConvertFromString($texto)
    $brocha.Freeze()
    return $brocha
}

function Set-ColorAcentoVentana {
    param($Ventana, [string]$Hex)
    $col = ConvertFrom-ColorHex $Hex
    $txt = Get-ColorTexto $col
    $hexTexto = '#{0:X2}{1:X2}{2:X2}' -f $txt.R, $txt.G, $txt.B
    $Ventana.Resources['Acento'] = New-BrochaWpf -Hex $Hex
    $Ventana.Resources['AcentoTexto'] = New-BrochaWpf -Hex $hexTexto
    $Ventana.Resources['AcentoSuave'] = New-BrochaWpf -Hex $Hex -Alfa 22
}

function Show-Interfaz {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

    $xml = [xml]$script:XamlInterfaz
    $ventana = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))

    # Referencias a los controles con nombre (x:Name)
    $c = @{}
    foreach ($nodo in $xml.SelectNodes('//*[@*[local-name()="Name"]]')) {
        $nombre = $nodo.Attributes | Where-Object { $_.LocalName -eq 'Name' } | Select-Object -First 1
        if ($null -ne $nombre) { $c[$nombre.Value] = $ventana.FindName($nombre.Value) }
    }

    $script:CajaRegistro = $c.txtLog
    $script:TextoEstado = $c.txtEstado
    $script:UltimaSalida = ''
    $script:LogoPendiente = $null     # logo de empresa elegido sin guardar ('' = volver a iniciales)
    $script:LogoAppRuta = ''          # logo de la app (elegido o extraído del EXE)
    $script:TipoActual = ''
    $script:Listo = $false
    $script:RutaInstalador = ''
    $script:PaginaActual = ''
    $script:CarpetaDestino = ''
    $script:BannerPendiente = $null   # banner elegido sin guardar ('' = volver al automático)

    # Ajustar al tamaño de la pantalla (portátiles con 768 px de alto)
    $area = [System.Windows.SystemParameters]::WorkArea
    if ($ventana.Height -gt ($area.Height - 10)) { $ventana.Height = [Math]::Max(600, $area.Height - 10) }
    if ($ventana.Width -gt ($area.Width - 10)) { $ventana.Width = [Math]::Max(960, $area.Width - 10) }

    # ---------- Navegación ----------
    $irA = {
        param([string]$Pagina)
        foreach ($par in @(@('Crear', $c.pagCrear, $c.navCrear), @('Paquetes', $c.pagPaquetes, $c.navPaquetes), @('Plantilla', $c.pagPlantilla, $c.navPlantilla), @('Ayuda', $c.pagAyuda, $c.navAyuda))) {
            if ($par[0] -eq $Pagina) {
                $par[1].Visibility = 'Visible'
                $par[2].Background = $ventana.Resources['AcentoSuave']
                $par[2].FontWeight = 'SemiBold'
            }
            else {
                $par[1].Visibility = 'Collapsed'
                $par[2].Background = [System.Windows.Media.Brushes]::Transparent
                $par[2].FontWeight = 'Normal'
            }
        }
        $script:PaginaActual = $Pagina
    }
    $c.navCrear.Add_Click({ & $irA 'Crear' })
    $c.navPaquetes.Add_Click({ & $irA 'Paquetes'; & $cargarPaquetes })
    $c.navPlantilla.Add_Click({ & $irA 'Plantilla' })
    $c.navAyuda.Add_Click({ & $irA 'Ayuda' })

    # ---------- Página: Paquetes creados ----------
    $abrirCarpeta = {
        param([string]$Ruta)
        if ($Ruta -and (Test-Path -LiteralPath $Ruta)) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Ruta) }
        else { [System.Windows.MessageBox]::Show("La carpeta ya no existe:`n$Ruta", 'Paquetes creados', 'OK', 'Warning') | Out-Null }
    }
    $cargarPaquetes = {
        $c.pnlPaquetes.Children.Clear()
        $filtro = $c.txtBuscarPaquete.Text.Trim()
        $todos = Get-HistorialPaquetes
        $mostrados = 0
        foreach ($p in $todos) {
            $textoBusqueda = ($p.NombrePaquete, $p.NombreApp, $p.Version, $p.Fabricante) -join ' '
            if ($filtro -and $textoBusqueda.IndexOf($filtro, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            $mostrados++
            $tarjeta = New-Object System.Windows.Controls.Border
            $tarjeta.Background = [System.Windows.Media.Brushes]::White
            $tarjeta.BorderBrush = New-BrochaWpf -Hex '#E4E7EC'
            $tarjeta.BorderThickness = 1
            $tarjeta.CornerRadius = 10
            $tarjeta.Padding = '16,12'
            $tarjeta.Margin = '0,0,0,10'
            $g = New-Object System.Windows.Controls.Grid
            $c1 = New-Object System.Windows.Controls.ColumnDefinition
            $c1.Width = [System.Windows.GridLength]::Auto
            $c2 = New-Object System.Windows.Controls.ColumnDefinition
            $c3 = New-Object System.Windows.Controls.ColumnDefinition
            $c3.Width = [System.Windows.GridLength]::Auto
            $g.ColumnDefinitions.Add($c1); $g.ColumnDefinitions.Add($c2); $g.ColumnDefinitions.Add($c3)

            $chip = New-Object System.Windows.Controls.Border
            $chip.Width = 44; $chip.Height = 44; $chip.CornerRadius = 8
            $chip.Background = $ventana.Resources['AcentoSuave']
            $chip.Margin = '0,0,14,0'
            $chipTexto = New-Object System.Windows.Controls.TextBlock
            if ($p.Tipo) { $chipTexto.Text = $p.Tipo } else { $chipTexto.Text = 'PKG' }
            $chipTexto.FontWeight = 'Bold'; $chipTexto.FontSize = 11
            $chipTexto.HorizontalAlignment = 'Center'; $chipTexto.VerticalAlignment = 'Center'
            $chipTexto.Foreground = $ventana.Resources['Acento']
            $chip.Child = $chipTexto

            $textos = New-Object System.Windows.Controls.StackPanel
            $textos.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($textos, 1)
            $t1 = New-Object System.Windows.Controls.TextBlock
            if ($p.NombreApp) { $t1.Text = ("{0} {1}" -f $p.NombreApp, $p.Version).Trim() } else { $t1.Text = $p.NombrePaquete }
            $t1.FontSize = 14; $t1.FontWeight = 'SemiBold'; $t1.TextTrimming = 'CharacterEllipsis'
            $t2 = New-Object System.Windows.Controls.TextBlock
            $detalle = @($p.Fecha)
            if ($p.Fabricante) { $detalle += $p.Fabricante }
            if ($p.Usuario) { $detalle += "por $($p.Usuario)" }
            if ($p.Intunewin -and (Test-Path -LiteralPath $p.Intunewin)) { $detalle += '.intunewin listo' }
            elseif ($p.ExisteSalida) { $detalle += 'sin .intunewin' }
            if (-not $p.ExisteFuente) { $detalle += 'carpeta fuente borrada' }
            $t2.Text = $detalle -join '  ·  '
            $t2.FontSize = 11.5; $t2.Foreground = New-BrochaWpf -Hex '#667085'; $t2.TextTrimming = 'CharacterEllipsis'
            $t3 = New-Object System.Windows.Controls.TextBlock
            $t3.Text = $p.NombrePaquete; $t3.FontSize = 11; $t3.FontFamily = 'Consolas'; $t3.Foreground = New-BrochaWpf -Hex '#98A2B3'
            [void]$textos.Children.Add($t1); [void]$textos.Children.Add($t2); [void]$textos.Children.Add($t3)

            $botones = New-Object System.Windows.Controls.StackPanel
            $botones.Orientation = 'Horizontal'; $botones.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($botones, 2)
            foreach ($b in @(@('📂  Salida', $p.RutaSalida, $p.ExisteSalida), @('🛠  Fuente', $p.RutaPaquete, $p.ExisteFuente))) {
                $boton = New-Object System.Windows.Controls.Button
                $boton.Style = $ventana.FindResource('BotonSecundario')
                $boton.Content = $b[0]
                $boton.Tag = $b[1]
                $boton.Margin = '8,0,0,0'
                $boton.IsEnabled = [bool]$b[2]
                $boton.ToolTip = [string]$b[1]
                $boton.Add_Click({ param($s, $e) & $abrirCarpeta ([string]$s.Tag) })
                [void]$botones.Children.Add($boton)
            }
            [void]$g.Children.Add($chip); [void]$g.Children.Add($textos); [void]$g.Children.Add($botones)
            $tarjeta.Child = $g
            [void]$c.pnlPaquetes.Children.Add($tarjeta)
        }
        if ($todos.Count -eq 0) {
            $vacio = New-Object System.Windows.Controls.TextBlock
            $vacio.Text = 'Todavía no hay paquetes. Crea el primero en «Crear paquete».'
            $vacio.Foreground = New-BrochaWpf -Hex '#667085'; $vacio.Margin = '4,20,0,0'
            [void]$c.pnlPaquetes.Children.Add($vacio)
        }
        $c.txtPaquetesResumen.Text = "Mostrando $mostrados de $($todos.Count) paquetes, del más reciente al más antiguo."
    }
    $c.btnRefrescarPaquetes.Add_Click({ & $cargarPaquetes })
    $c.txtBuscarPaquete.Add_TextChanged({ & $cargarPaquetes })

    # ---------- Versión web ----------
    $c.btnModoWeb.Add_Click({
            try {
                $argumentos = '-NoProfile -ExecutionPolicy Bypass -STA -File "{0}" -ModoWeb' -f $script:RutaEsteScript
                Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $argumentos -WindowStyle Minimized
                Write-Registro -Mensaje 'Abriendo la versión web en el navegador (se abre una consola minimizada: no la cierres mientras la uses).' -Tipo OK
            }
            catch {
                Write-Registro -Mensaje "No se pudo abrir la versión web: $($_.Exception.Message)" -Tipo Error
            }
        })

    # ---------- Etiquetas de deslizadores ----------
    $textoAplazar = { param([int]$v) if ($v -eq 0) { 'No permitir' } elseif ($v -eq 1) { '1 vez' } else { "$v veces" } }
    $textoDias = { param([int]$v) if ($v -eq 0) { 'Sin límite' } elseif ($v -eq 1) { '1 día' } else { "$v días" } }
    $textoCuenta = { param([int]$v) if ($v -eq 0) { 'Sin cuenta' } else { "$v min" } }

    # ---------- Vista previa de la ventana (maqueta) ----------
    $actualizarVista = {
        if (-not $script:Listo) { return }
        $cfg = Get-ConfigCorporativa
        $empresa = [string]$c.txtEmpresa.Text
        if (-not $empresa) { $empresa = [string]$cfg.Empresa }
        $nombre = $c.txtNombre.Text.Trim(); if (-not $nombre) { $nombre = 'Nombre de la app' }
        $version = $c.txtVersion.Text.Trim()
        $c.mockTitulo.Text = ("$nombre $version").Trim()
        $c.mockSubtitulo.Text = "$empresa - Instalación de la aplicación"

        # Logo
        $rutaLogo = ''
        if ($c.chkLogoVentanas.IsChecked -and $script:LogoAppRuta) { $rutaLogo = $script:LogoAppRuta }
        else { $rutaLogo = Get-RutaLogoCorporativo }
        $c.mockLogo.Source = Get-ImagenWpf $rutaLogo
        $c.mockInicialesTexto.Text = Get-Iniciales $empresa
        $c.mockIniciales.Visibility = if ($null -eq $c.mockLogo.Source) { 'Visible' } else { 'Collapsed' }

        $procesos = Get-ListaProcesos -Procesos $c.txtProcesos.Text -Descripcion $nombre
        $veces = [int]$c.sldAplazar.Value
        $dias = [int]$c.sldDias.Value
        $cuenta = [int]$c.sldCuenta.Value

        # Textos estándar de PSADT en español; si el usuario escribe su texto, lo reemplaza.
        if ($c.txtBienvenida.Text.Trim()) {
            Set-TextoFormateado -Bloque $c.mockTexto -Texto $c.txtBienvenida.Text.Trim()
        }
        elseif ($procesos.Count -gt 0) {
            Set-TextoFormateado -Bloque $c.mockTexto -Texto 'Por favor, guarde su trabajo antes de continuar ya que las siguientes aplicaciones se cerrarán automáticamente.'
        }
        else {
            Set-TextoFormateado -Bloque $c.mockTexto -Texto 'Por favor, seleccione Instalar para continuar con la instalación.'
        }
        $c.mockMensaje.Visibility = 'Collapsed'

        $c.mockLista.Children.Clear()
        foreach ($p in $procesos) {
            $fila = New-Object System.Windows.Controls.StackPanel
            $fila.Orientation = 'Horizontal'
            $fila.Margin = '0,4,0,4'
            $icono = New-Object System.Windows.Controls.Border
            $icono.Width = 22; $icono.Height = 22; $icono.CornerRadius = 4
            $icono.Background = $ventana.Resources['AcentoSuave']
            $letra = New-Object System.Windows.Controls.TextBlock
            $letra.Text = ([string]$p.Descripcion).Substring(0, 1).ToUpper()
            $letra.HorizontalAlignment = 'Center'; $letra.VerticalAlignment = 'Center'; $letra.FontWeight = 'Bold'; $letra.FontSize = 11
            $icono.Child = $letra
            $textos = New-Object System.Windows.Controls.StackPanel
            $textos.Margin = '8,0,0,0'
            $t1 = New-Object System.Windows.Controls.TextBlock
            $t1.Text = [string]$p.Descripcion; $t1.FontSize = 11.5; $t1.FontWeight = 'SemiBold'
            $t2 = New-Object System.Windows.Controls.TextBlock
            $t2.Text = ([string]$p.Nombre + '.exe'); $t2.FontSize = 10; $t2.Foreground = $ventana.Resources['Acento']
            [void]$textos.Children.Add($t1); [void]$textos.Children.Add($t2)
            [void]$fila.Children.Add($icono); [void]$fila.Children.Add($textos)
            [void]$c.mockLista.Children.Add($fila)
        }
        $c.mockListaBorde.Visibility = if ($procesos.Count -gt 0) { 'Visible' } else { 'Collapsed' }

        $c.mockFilas.Children.Clear()
        $agregarFila = {
            param([string]$icono, [string]$texto, [string]$valor, [bool]$resaltar)
            $g = New-Object System.Windows.Controls.Grid
            $g.Margin = '0,6,0,0'
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = [System.Windows.GridLength]::Auto
            $g.ColumnDefinitions.Add($col1); $g.ColumnDefinitions.Add($col2)
            $a = New-Object System.Windows.Controls.TextBlock
            $a.Text = "$icono  $texto"; $a.FontSize = 11.5; $a.Foreground = [System.Windows.Media.Brushes]::DimGray
            $b = New-Object System.Windows.Controls.TextBlock
            $b.Text = $valor; $b.FontSize = 11.5; $b.FontWeight = 'SemiBold'
            if ($resaltar) { $b.Foreground = [System.Windows.Media.Brushes]::DarkGoldenrod }
            [System.Windows.Controls.Grid]::SetColumn($b, 1)
            [void]$g.Children.Add($a); [void]$g.Children.Add($b)
            [void]$c.mockFilas.Children.Add($g)
        }
        if ($cuenta -gt 0 -and ($procesos.Count -gt 0 -or $veces -gt 0)) { & $agregarFila '⏱' 'Cuenta regresiva de inicio automático' ('0h {0}m 0s' -f $cuenta) $false }
        if ($veces -gt 0) { & $agregarFila '💬' 'Aplazamientos restantes' ([string]$veces) $false }
        if ($veces -gt 0 -and $dias -gt 0) { & $agregarFila '📅' 'Fecha límite de aplazamiento' ((Get-Date).AddDays($dias).ToString('dd/MM/yyyy hh:mm tt')) $true }

        if ($c.txtBotonInstalar.Text.Trim()) { $c.mockBtnInstalarTexto.Text = $c.txtBotonInstalar.Text.Trim() }
        elseif ($procesos.Count -gt 0) { $c.mockBtnInstalarTexto.Text = 'Cerrar Aplicaciones e Instalar' }
        else { $c.mockBtnInstalarTexto.Text = 'Instalar' }
        if ($c.txtBotonAplazar.Text.Trim()) { $c.mockBtnAplazarTexto.Text = $c.txtBotonAplazar.Text.Trim() }
        else { $c.mockBtnAplazarTexto.Text = 'Aplazar' }
        $c.mockBtnAplazar.Visibility = if ($veces -gt 0) { 'Visible' } else { 'Collapsed' }
        if ($veces -gt 0) { [System.Windows.Controls.Grid]::SetColumnSpan($c.mockBtnInstalar, 1) }
        else { [System.Windows.Controls.Grid]::SetColumnSpan($c.mockBtnInstalar, 3) }

        # Aviso de cuándo se ven las ventanas
        if ($procesos.Count -gt 0 -and -not $c.chkSiempre.IsChecked) {
            $c.txtAvisoSilencio.Text = 'Esta ventana solo aparece si alguna de esas apps está ABIERTA. Si están cerradas, se instala en silencio sin molestar al usuario.'
        }
        elseif ($procesos.Count -eq 0 -and $veces -eq 0 -and $cuenta -eq 0) {
            $c.txtAvisoSilencio.Text = 'Sin apps a cerrar ni aplazamiento: el usuario solo verá la barra de progreso.'
        }
        else {
            $c.txtAvisoSilencio.Text = 'Esta ventana aparece cada vez que haya un usuario conectado en el equipo.'
        }
        & $validarFormulario
    }

    # ---------- Validación en vivo y avance ----------
    $marcarPaso = {
        param($Bloque, [bool]$Ok, [string]$Texto)
        if ($Ok) {
            $Bloque.Text = "✔  $Texto"
            $Bloque.Foreground = New-BrochaWpf -Hex '#039855'
            $Bloque.FontWeight = 'SemiBold'
        }
        else {
            $Bloque.Text = "○  $Texto"
            $Bloque.Foreground = New-BrochaWpf -Hex '#98A2B3'
            $Bloque.FontWeight = 'Normal'
        }
    }
    $validarFormulario = {
        if (-not $script:Listo) { return }
        $hayInstalador = [bool]$script:RutaInstalador
        $r = Get-ProblemasDatos -Datos (& $obtenerDatos) -Tipo $script:TipoActual
        $mapa = @{
            NombreApp             = @($c.txtNombre, $c.errNombre)
            Version               = @($c.txtVersion, $c.errVersion)
            DesinstaladorExe      = @($c.txtDes, $c.errDes)
            ArgumentosInstalacion = @($c.txtArgs, $c.errArgs)
        }
        foreach ($k in $mapa.Keys) { $mapa[$k][0].Tag = $null; $mapa[$k][1].Visibility = 'Collapsed' }
        $errores = @($r.Errores | Where-Object { $_.Campo -ne 'Instalador' })
        if ($hayInstalador) {
            foreach ($problema in @(@($r.Avisos | ForEach-Object { @{ P = $_; Nivel = 'Aviso' } }) + @($errores | ForEach-Object { @{ P = $_; Nivel = 'Error' } }))) {
                $campo = $problema.P.Campo
                if (-not $mapa.ContainsKey($campo)) { continue }
                $mapa[$campo][0].Tag = $problema.Nivel
                $mapa[$campo][1].Text = $problema.P.Texto
                $mapa[$campo][1].Visibility = 'Visible'
            }
        }
        $okInfo = $hayInstalador -and (@($errores | Where-Object { $_.Campo -in @('NombreApp', 'Version') }).Count -eq 0)
        $okInstalacion = $hayInstalador -and (@($errores | Where-Object { $_.Campo -eq 'DesinstaladorExe' }).Count -eq 0)
        & $marcarPaso $c.pasoInstalador $hayInstalador '1. Instalador'
        & $marcarPaso $c.pasoInfo $okInfo '2. Información de la app'
        & $marcarPaso $c.pasoInstalacion $okInstalacion '3. Cómo se instala'
        & $marcarPaso $c.pasoVentanas $hayInstalador '4. Ventanas del usuario'

        if (-not $hayInstalador) {
            $c.txtResumen.Text = 'Arrastra un instalador para empezar.'
            $c.bdResumen.Background = New-BrochaWpf -Hex '#F2F4F7'
            $c.txtResumen.Foreground = New-BrochaWpf -Hex '#344054'
        }
        elseif ($errores.Count -gt 0) {
            $c.txtResumen.Text = 'Antes de crear: ' + (($errores | ForEach-Object { $_.Texto }) -join ' ')
            $c.bdResumen.Background = New-BrochaWpf -Hex '#FEF3F2'
            $c.txtResumen.Foreground = New-BrochaWpf -Hex '#B42318'
        }
        elseif (@($r.Avisos).Count -gt 0) {
            $c.txtResumen.Text = 'Listo para crear. Ten en cuenta: ' + ((@($r.Avisos) | ForEach-Object { $_.Texto }) -join ' ')
            $c.bdResumen.Background = New-BrochaWpf -Hex '#FFFAEB'
            $c.txtResumen.Foreground = New-BrochaWpf -Hex '#B54708'
        }
        else {
            $c.txtResumen.Text = '✔ Todo listo para crear el paquete.'
            $c.bdResumen.Background = New-BrochaWpf -Hex '#ECFDF3'
            $c.txtResumen.Foreground = New-BrochaWpf -Hex '#027A48'
        }
    }

    # ---------- Tipo de instalador ----------
    $aplicarModo = {
        param([string]$Tipo)
        $script:TipoActual = $Tipo
        $visible = [System.Windows.Visibility]::Visible
        $oculto = [System.Windows.Visibility]::Collapsed
        $c.bdTipo.Visibility = if ($Tipo) { $visible } else { $oculto }
        $c.txtTipo.Text = $Tipo
        $c.chkCarpeta.IsEnabled = $true
        switch ($Tipo) {
            'MSI' {
                $c.txtSubInstalacion.Text = 'MSI: no hay que escribir nada. Se instala, desinstala y detecta con el propio MSI.'
                $c.filaArgs.Visibility = $visible; $c.filaDes.Visibility = $oculto; $c.filaArgsDes.Visibility = $oculto; $c.filaDet.Visibility = $oculto
                $c.lblArgs.Text = 'Propiedades MSI (opcional)'
                $c.lblAyudaArgs.Text = 'Ej: SERVIDOR=srv01 LICENCIA=xyz. El modo silencioso (/qn) ya lo pone PSADT.'
            }
            'EXE' {
                $c.txtSubInstalacion.Text = 'EXE: indica cómo se instala en silencio y dónde queda su desinstalador.'
                $c.filaArgs.Visibility = $visible; $c.filaDes.Visibility = $visible; $c.filaArgsDes.Visibility = $visible; $c.filaDet.Visibility = $visible
                $c.lblArgs.Text = 'Argumentos silenciosos (obligatorio)'
                $c.lblAyudaArgs.Text = 'Ej: /S   /silent   /quiet   /VERYSILENT /NORESTART  (prueba setup.exe /? para verlos)'
                $c.lblDes.Text = 'Desinstalador en el equipo (obligatorio)'
                $c.lblAyudaDes.Text = 'Ruta después de instalar. Ej: $env:ProgramFiles\MiApp\uninstall.exe'
                $c.btnDes.Visibility = $oculto
            }
            'PS1' {
                $c.txtSubInstalacion.Text = 'Script PS1: se copia toda su carpeta. Tu script debe terminar con «exit 0» si todo salió bien.'
                $c.filaArgs.Visibility = $visible; $c.filaDes.Visibility = $visible; $c.filaArgsDes.Visibility = $visible; $c.filaDet.Visibility = $oculto
                $c.lblArgs.Text = 'Argumentos del script (opcional)'
                $c.lblAyudaArgs.Text = 'Se pasan al script tal cual. Ej: -Modo Completo'
                $c.lblDes.Text = 'Script de desinstalación (opcional)'
                $c.lblAyudaDes.Text = 'Debe estar en la misma carpeta. Si lo dejas vacío, al desinstalar solo se quita la marca de registro.'
                $c.btnDes.Visibility = $visible
                $c.chkCarpeta.IsChecked = $true; $c.chkCarpeta.IsEnabled = $false
            }
            default {
                $c.txtSubInstalacion.Text = 'Primero elige un instalador.'
                $c.filaArgs.Visibility = $oculto; $c.filaDes.Visibility = $oculto; $c.filaArgsDes.Visibility = $oculto; $c.filaDet.Visibility = $oculto
            }
        }
    }

    $mostrarLogoApp = {
        param([string]$Ruta, [string]$Info)
        if ($Ruta -and (Test-Path -LiteralPath $Ruta)) {
            $script:LogoAppRuta = $Ruta
            $c.imgLogoApp.Source = Get-ImagenWpf $Ruta
            if ($Info) { $c.txtLogoAppInfo.Text = $Info } else { $c.txtLogoAppInfo.Text = [System.IO.Path]::GetFileName($Ruta) }
            $c.chkLogoVentanas.IsEnabled = $true
        }
        else {
            $script:LogoAppRuta = ''
            $c.imgLogoApp.Source = $null
            $c.txtLogoAppInfo.Text = 'Sin logo: se generan las iniciales con el color de la empresa.'
            $c.chkLogoVentanas.IsChecked = $false
            $c.chkLogoVentanas.IsEnabled = $false
        }
        $n = $c.txtNombre.Text.Trim(); if (-not $n) { $n = 'App' }
        $c.txtLogoAppIniciales.Text = Get-Iniciales $n
        & $actualizarVista
    }

    $cargarInstalador = {
        param([string]$Ruta)
        try {
            $info = Get-InfoInstalador -Ruta $Ruta
            $cfg = Get-ConfigCorporativa
            $script:RutaInstalador = $Ruta
            $c.txtZonaIcono.Text = '✔'
            $c.txtZonaTitulo.Text = [System.IO.Path]::GetFileName($Ruta)
            $c.txtZonaDetalle.Text = (Split-Path -Path $Ruta -Parent) + '   (clic para cambiar)'
            $c.txtNombre.Text = $info.Nombre
            $c.txtVersion.Text = $info.Version
            $c.txtFabricante.Text = $info.Fabricante
            if ([string]::IsNullOrWhiteSpace($c.txtFabricante.Text)) { $c.txtFabricante.Text = [string]$cfg.Empresa }
            $c.txtDeteccion.Text = $info.Nombre
            $c.txtArgs.Text = ''
            $c.txtDes.Text = $info.Desinstalador
            $c.txtArgsDes.Text = ''
            $c.chkCarpeta.IsChecked = $false
            & $aplicarModo $info.Tipo
            if ($info.Tipo -eq 'MSI' -and $info.CodigoProducto) { $c.txtTipoDetalle.Text = "ProductCode $($info.CodigoProducto)" }
            elseif ($info.Tipo -eq 'PS1' -and $info.Desinstalador) { $c.txtTipoDetalle.Text = 'Script de desinstalación detectado: ' + [System.IO.Path]::GetFileName($info.Desinstalador) }
            else { $c.txtTipoDetalle.Text = '' }

            & $mostrarLogoApp '' ''
            if ($info.Tipo -eq 'EXE') {
                $icono = Get-IconoDeExe -Ruta $Ruta
                if ($null -ne $icono) {
                    $rutaIcono = Join-Path $env:TEMP ('FabricaIcono_' + [guid]::NewGuid().ToString('N') + '.png')
                    try { $icono.Save($rutaIcono, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $icono.Dispose() }
                    & $mostrarLogoApp $rutaIcono 'Ícono extraído del EXE. Puedes cambiarlo.'
                }
            }
            Write-Registro -Mensaje "Instalador leído: $($info.Nombre) $($info.Version) [$($info.Tipo)]" -Tipo OK
            if ([string]::IsNullOrWhiteSpace($info.Version)) { Write-Registro -Mensaje 'El instalador no trae versión: escríbela a mano.' -Tipo Aviso }
            & $actualizarVista
        }
        catch {
            & $aplicarModo ''
            Write-Registro -Mensaje $_.Exception.Message -Tipo Error
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'No se pudo leer el instalador', 'OK', 'Error') | Out-Null
        }
    }

    $obtenerDatos = {
        $veces = [int]$c.sldAplazar.Value
        return @{
            Instalador               = [string]$script:RutaInstalador
            Fabricante               = $c.txtFabricante.Text.Trim()
            NombreApp                = $c.txtNombre.Text.Trim()
            Version                  = $c.txtVersion.Text.Trim()
            Descripcion              = $c.txtDescripcion.Text.Trim()
            Arquitectura             = [string]$c.cmbArq.SelectedItem.Content
            ArgumentosInstalacion    = $c.txtArgs.Text.Trim()
            DesinstaladorExe         = $c.txtDes.Text.Trim()
            ArgumentosDesinstalacion = $c.txtArgsDes.Text.Trim()
            NombreDeteccion          = $c.txtDeteccion.Text.Trim()
            ProcesosACerrar          = $c.txtProcesos.Text.Trim()
            SiemprePreguntar         = [bool]$c.chkSiempre.IsChecked
            PermitirAplazar          = ($veces -gt 0)
            VecesAplazar             = [Math]::Max(1, $veces)
            DiasLimite               = [int]$c.sldDias.Value
            CuentaRegresivaMin       = [int]$c.sldCuenta.Value
            MensajeBienvenida        = $c.txtBienvenida.Text.Trim()
            TextoBotonInstalar       = $c.txtBotonInstalar.Text.Trim()
            TextoBotonAplazar        = $c.txtBotonAplazar.Text.Trim()
            MensajeFinal             = $c.txtMensaje.Text.Trim()
            LogoApp                  = $script:LogoAppRuta
            LogoAppEnVentanas        = [bool]$c.chkLogoVentanas.IsChecked
            CopiarCarpetaCompleta    = [bool]$c.chkCarpeta.IsChecked
            GenerarIntunewin         = [bool]$c.chkIntunewin.IsChecked
            CarpetaDestino           = [string]$script:CarpetaDestino
            IncluirFuente            = [bool]$c.chkCopiarFuente.IsChecked
        }
    }

    # ---------- Carpeta destino ----------
    $mostrarDestino = {
        if ($script:CarpetaDestino) {
            $c.txtDestino.Text = $script:CarpetaDestino
            $c.txtDestino.ToolTip = $script:CarpetaDestino
        }
        else {
            $c.txtDestino.Text = "$($script:CarpetaSalida)  (por defecto)"
            $c.txtDestino.ToolTip = $script:CarpetaSalida
        }
    }
    $guardarDestino = {
        try {
            $cfgD = Get-ConfigCorporativa
            $cfgD.CarpetaDestino = [string]$script:CarpetaDestino
            Save-ConfigCorporativa -Cfg $cfgD
        }
        catch { Write-Registro -Mensaje "No se pudo recordar la carpeta destino: $($_.Exception.Message)" -Tipo Aviso }
    }
    $c.btnDestino.Add_Click({
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = 'Elige la carpeta donde quedará el paquete listo (.intunewin, ficha, detección y logo)'
            $dlg.ShowNewFolderButton = $true
            if ($script:CarpetaDestino -and (Test-Path -LiteralPath $script:CarpetaDestino)) { $dlg.SelectedPath = $script:CarpetaDestino }
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $script:CarpetaDestino = $dlg.SelectedPath
                & $mostrarDestino
                & $guardarDestino
                Write-Registro -Mensaje "Los paquetes quedarán en: $($script:CarpetaDestino)" -Tipo OK
            }
        })
    $c.btnDestinoDefecto.Add_Click({
            $script:CarpetaDestino = ''
            & $mostrarDestino
            & $guardarDestino
            Write-Registro -Mensaje "Los paquetes quedarán en la carpeta Salida de la Fábrica." -Tipo OK
        })

    # ---------- Plantilla corporativa ----------
    $marcarPendiente = { if ($script:Listo) { $c.bdPendiente.Visibility = 'Visible' } }

    $cargarPlantilla = {
        $script:Listo = $false
        $cfg = Get-ConfigCorporativa
        $c.txtEmpresa.Text = [string]$cfg.Empresa
        $c.txtColor.Text = [string]$cfg.ColorAcento
        Set-ColorAcentoVentana -Ventana $ventana -Hex ([string]$cfg.ColorAcento)
        $c.cmbEstilo.SelectedIndex = if ([string]$cfg.EstiloDialogo -eq 'Classic') { 1 } else { 0 }
        $c.cmbIdioma.SelectedIndex = switch ([string]$cfg.Idioma) { 'es' { 0 } 'en' { 1 } default { 2 } }
        $c.txtLogs.Text = [string]$cfg.RutaLogs
        $c.sldAplazarDef.Value = [Math]::Min(10, [int]$cfg.VecesAplazar)
        $c.sldDiasDef.Value = [Math]::Min(30, [int]$cfg.DiasLimite)
        $c.sldCuentaDef.Value = [Math]::Min(60, [int]$cfg.CuentaRegresivaMin)
        $c.txtBienvenidaDef.Text = [string]$cfg.MensajeBienvenida
        $c.txtMensajeDef.Text = [string]$cfg.MensajeFinal
        $c.txtBotonInstalarDef.Text = [string]$cfg.TextoBotonInstalar
        $c.txtBotonAplazarDef.Text = [string]$cfg.TextoBotonAplazar
        $script:LogoPendiente = $null
        $rutaLogoEmp = ''
        if ([string]$cfg.Logo) { $rutaLogoEmp = Join-Path $script:Recursos ([string]$cfg.Logo) }
        $c.imgLogoEmp.Source = Get-ImagenWpf $rutaLogoEmp
        if ([string]$cfg.Logo) { $c.txtLogoEmpInfo.Text = 'Logo guardado: ' + [string]$cfg.Logo }
        else { $c.txtLogoEmpInfo.Text = 'Sin logo: se usan las iniciales de la empresa.' }

        $script:BannerPendiente = $null
        if ([string]$cfg.Banner) {
            $c.imgBannerEmp.Source = Get-ImagenWpf (Join-Path $script:Recursos ([string]$cfg.Banner))
            $c.txtBannerEmpInfo.Text = 'Banner propio: ' + [string]$cfg.Banner
        }
        else {
            $c.imgBannerEmp.Source = Get-ImagenWpf (Get-RutaBannerCorporativo)
            $c.txtBannerEmpInfo.Text = 'Generado automáticamente con el logo, el color y el nombre de la empresa.'
        }

        $rutaIcono = Get-RutaLogoCorporativo
        $c.imgIconoPrev.Source = Get-ImagenWpf $rutaIcono
        $c.imgEncabezado.Source = Get-ImagenWpf $rutaIcono
        $c.imgBannerPrev.Source = Get-ImagenWpf (Get-RutaBannerCorporativo)
        $c.bdPendiente.Visibility = 'Collapsed'

        # Valores por defecto en "Crear paquete"
        $c.sldAplazar.Value = [Math]::Min(10, [int]$cfg.VecesAplazar)
        $c.sldDias.Value = [Math]::Min(30, [int]$cfg.DiasLimite)
        $c.sldCuenta.Value = [Math]::Min(60, [int]$cfg.CuentaRegresivaMin)
        $c.txtBienvenida.Text = [string]$cfg.MensajeBienvenida
        $c.txtBotonInstalar.Text = [string]$cfg.TextoBotonInstalar
        $c.txtBotonAplazar.Text = [string]$cfg.TextoBotonAplazar
        $c.txtMensaje.Text = [string]$cfg.MensajeFinal
        $script:CarpetaDestino = [string]$cfg.CarpetaDestino
        & $mostrarDestino
        $script:Listo = $true
        & $actualizarVista
    }

    # Paleta de colores rápida
    foreach ($hex in @('#FDDA24', '#2C2A29', '#0F4C81', '#0078D4', '#00A651', '#E4002B', '#6941C6')) {
        $muestra = New-Object System.Windows.Controls.Button
        $muestra.Width = 24; $muestra.Height = 24; $muestra.Margin = '0,0,6,0'; $muestra.Cursor = 'Hand'
        $muestra.ToolTip = $hex
        $muestra.Tag = $hex
        $muestra.Background = New-BrochaWpf -Hex $hex
        $muestra.BorderThickness = 1
        $muestra.Add_Click({ param($s, $e) $c.txtColor.Text = [string]$s.Tag })
        [void]$c.pnlPaleta.Children.Add($muestra)
    }

    # ---------- Eventos: Crear paquete ----------
    $elegirInstalador = {
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = 'Instaladores (*.msi;*.exe;*.ps1)|*.msi;*.exe;*.ps1'
        $dlg.Title = 'Selecciona el instalador'
        if ($dlg.ShowDialog($ventana) -eq $true) { & $cargarInstalador $dlg.FileName }
    }
    $c.zonaArrastre.Add_MouseLeftButtonUp({ & $elegirInstalador })
    $c.zonaArrastre.Add_DragOver({
            param($s, $e)
            if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { $e.Effects = [System.Windows.DragDropEffects]::Copy }
            else { $e.Effects = [System.Windows.DragDropEffects]::None }
            $e.Handled = $true
        })
    $c.zonaArrastre.Add_Drop({
            param($s, $e)
            $archivos = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
            if ($archivos -and $archivos.Count -gt 0) { & $cargarInstalador ([string]$archivos[0]) }
        })

    $c.btnDes.Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Filter = 'Script PowerShell (*.ps1)|*.ps1'
            $dlg.Title = 'Selecciona el script de desinstalación'
            if ($script:RutaInstalador) { $dlg.InitialDirectory = Split-Path -Path $script:RutaInstalador -Parent }
            if ($dlg.ShowDialog($ventana) -eq $true) { $c.txtDes.Text = $dlg.FileName }
        })

    $c.btnLogoApp.Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Filter = 'Imágenes (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico)|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico'
            $dlg.Title = 'Selecciona el logo de la aplicación'
            if ($dlg.ShowDialog($ventana) -eq $true) {
                if ($null -eq (Get-ImagenWpf $dlg.FileName)) {
                    [System.Windows.MessageBox]::Show('No se pudo abrir la imagen.', 'Logo de la app', 'OK', 'Warning') | Out-Null
                    return
                }
                & $mostrarLogoApp $dlg.FileName ''
            }
        })
    $c.btnQuitarLogoApp.Add_Click({ & $mostrarLogoApp '' '' })

    foreach ($ctl in @($c.txtNombre, $c.txtVersion, $c.txtProcesos, $c.txtBienvenida, $c.txtBotonInstalar, $c.txtBotonAplazar)) { $ctl.Add_TextChanged({ & $actualizarVista }) }
    $c.txtNombre.Add_TextChanged({ if (-not $script:LogoAppRuta) { $n = $c.txtNombre.Text.Trim(); if (-not $n) { $n = 'App' }; $c.txtLogoAppIniciales.Text = Get-Iniciales $n } })
    foreach ($ctl in @($c.txtFabricante, $c.txtDescripcion, $c.txtArgs, $c.txtDes)) { $ctl.Add_TextChanged({ & $validarFormulario }) }
    foreach ($ctl in @($c.chkLogoVentanas, $c.chkSiempre)) { $ctl.Add_Checked({ & $actualizarVista }); $ctl.Add_Unchecked({ & $actualizarVista }) }
    $c.sldAplazar.Add_ValueChanged({
            $c.lblAplazar.Text = & $textoAplazar ([int]$c.sldAplazar.Value)
            $c.sldDias.IsEnabled = ([int]$c.sldAplazar.Value -gt 0)
            & $actualizarVista
        })
    $c.sldDias.Add_ValueChanged({ $c.lblDias.Text = & $textoDias ([int]$c.sldDias.Value); & $actualizarVista })
    $c.sldCuenta.Add_ValueChanged({ $c.lblCuenta.Text = & $textoCuenta ([int]$c.sldCuenta.Value); & $actualizarVista })

    $c.btnCrear.Add_Click({
            if (-not $script:RutaInstalador) {
                [System.Windows.MessageBox]::Show('Primero arrastra o elige el instalador.', 'Falta el instalador', 'OK', 'Warning') | Out-Null
                return
            }
            $problemas = Get-ProblemasDatos -Datos (& $obtenerDatos) -Tipo $script:TipoActual
            if (@($problemas.Errores).Count -gt 0) {
                & $validarFormulario
                $lista = (@($problemas.Errores) | ForEach-Object { '• ' + $_.Texto }) -join "`n"
                [System.Windows.MessageBox]::Show("Corrige esto antes de crear el paquete (está marcado en rojo):`n`n$lista", 'Faltan datos', 'OK', 'Warning') | Out-Null
                return
            }
            $c.btnCrear.IsEnabled = $false
            $c.btnCrear.Content = 'Creando paquete...'
            $ventana.Cursor = [System.Windows.Input.Cursors]::Wait
            $c.expRegistro.IsExpanded = $true
            try {
                $resultado = New-PaqueteCorporativo -Datos (& $obtenerDatos)
                $script:UltimaSalida = $resultado.RutaSalida
                $c.btnAbrir.IsEnabled = $true
                $r = [System.Windows.MessageBox]::Show("Paquete creado en:`n$($resultado.RutaSalida)`n`nAhí están el .intunewin, FichaIntune.txt, Deteccion.ps1, Descripcion.txt y Logo_Intune.png.`n`n¿Abrir la carpeta ahora?", 'Listo', 'YesNo', 'Information')
                if ($r -eq 'Yes') { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:UltimaSalida) }
            }
            catch {
                Write-Registro -Mensaje $_.Exception.Message -Tipo Error
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'No se pudo crear el paquete', 'OK', 'Error') | Out-Null
            }
            finally {
                $c.btnCrear.IsEnabled = $true
                $c.btnCrear.Content = 'Crear paquete'
                $ventana.Cursor = $null
            }
        })

    $c.btnVista.Add_Click({
            $c.btnVista.IsEnabled = $false
            $ventana.Cursor = [System.Windows.Input.Cursors]::Wait
            try { Start-VistaPrevia -Datos (& $obtenerDatos) }
            catch {
                Write-Registro -Mensaje $_.Exception.Message -Tipo Error
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'No se pudo abrir la vista previa', 'OK', 'Error') | Out-Null
            }
            finally {
                $c.btnVista.IsEnabled = $true
                $ventana.Cursor = $null
            }
        })

    $c.btnAbrir.Add_Click({
            if ($script:UltimaSalida -and (Test-Path -LiteralPath $script:UltimaSalida)) {
                Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:UltimaSalida)
            }
        })

    # ---------- Eventos: Plantilla ----------
    $c.txtColor.Add_TextChanged({
            $hex = $c.txtColor.Text.Trim()
            if ($hex.TrimStart('#') -match '^[0-9A-Fa-f]{6}$') {
                try {
                    $c.bdColor.Background = New-BrochaWpf -Hex $hex
                    Set-ColorAcentoVentana -Ventana $ventana -Hex $hex
                    if ($script:PaginaActual) { & $irA $script:PaginaActual }
                    & $marcarPendiente
                    & $actualizarVista
                }
                catch { Write-Registro -Mensaje "No se pudo aplicar el color: $($_.Exception.Message)" -Tipo Aviso }
            }
        })
    $c.btnColor.Add_Click({
            $dlg = New-Object System.Windows.Forms.ColorDialog
            $dlg.FullOpen = $true
            $dlg.Color = ConvertFrom-ColorHex $c.txtColor.Text
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $c.txtColor.Text = '#{0:X2}{1:X2}{2:X2}' -f $dlg.Color.R, $dlg.Color.G, $dlg.Color.B
            }
        })
    $c.btnLogo.Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Filter = 'Imágenes (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif'
            $dlg.Title = 'Selecciona el logo de la empresa'
            if ($dlg.ShowDialog($ventana) -eq $true) {
                $img = Get-ImagenWpf $dlg.FileName
                if ($null -eq $img) {
                    [System.Windows.MessageBox]::Show('No se pudo abrir la imagen.', 'Logo', 'OK', 'Warning') | Out-Null
                    return
                }
                $script:LogoPendiente = $dlg.FileName
                $c.imgLogoEmp.Source = $img
                $c.txtLogoEmpInfo.Text = [System.IO.Path]::GetFileName($dlg.FileName) + ' (sin guardar)'
                & $marcarPendiente
            }
        })
    $c.btnBanner.Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Filter = 'Imágenes (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp'
            $dlg.Title = 'Selecciona el banner (recomendado PNG de 900 x 125 px)'
            if ($dlg.ShowDialog($ventana) -eq $true) {
                $img = Get-ImagenWpf $dlg.FileName
                if ($null -eq $img) {
                    [System.Windows.MessageBox]::Show('No se pudo abrir la imagen.', 'Banner', 'OK', 'Warning') | Out-Null
                    return
                }
                $script:BannerPendiente = $dlg.FileName
                $c.imgBannerEmp.Source = $img
                $proporcion = $img.PixelWidth / [double]$img.PixelHeight
                $info = '{0} ({1} x {2} px) - sin guardar' -f [System.IO.Path]::GetFileName($dlg.FileName), $img.PixelWidth, $img.PixelHeight
                if ([Math]::Abs($proporcion - 7.2) / 7.2 -gt 0.15) { $info += '. Proporción distinta a 7,2:1: se centrará sobre fondo blanco.' }
                $c.txtBannerEmpInfo.Text = $info
                & $marcarPendiente
            }
        })
    $c.btnQuitarBanner.Add_Click({
            $script:BannerPendiente = ''
            $c.imgBannerEmp.Source = $null
            $c.txtBannerEmpInfo.Text = 'Se generará automáticamente (sin guardar)'
            & $marcarPendiente
        })
    $c.btnQuitarLogo.Add_Click({
            $script:LogoPendiente = ''
            $c.imgLogoEmp.Source = $null
            $c.txtLogoEmpInfo.Text = 'Se usarán las iniciales (sin guardar)'
            & $marcarPendiente
        })
    foreach ($ctl in @($c.txtEmpresa, $c.txtLogs, $c.txtBienvenidaDef, $c.txtMensajeDef, $c.txtBotonInstalarDef, $c.txtBotonAplazarDef)) { $ctl.Add_TextChanged({ & $marcarPendiente }) }
    $c.txtEmpresa.Add_TextChanged({ & $actualizarVista })
    foreach ($ctl in @($c.cmbEstilo, $c.cmbIdioma)) { $ctl.Add_SelectionChanged({ & $marcarPendiente }) }
    $c.sldAplazarDef.Add_ValueChanged({ $c.lblAplazarDef.Text = & $textoAplazar ([int]$c.sldAplazarDef.Value); & $marcarPendiente })
    $c.sldDiasDef.Add_ValueChanged({ $c.lblDiasDef.Text = & $textoDias ([int]$c.sldDiasDef.Value); & $marcarPendiente })
    $c.sldCuentaDef.Add_ValueChanged({ $c.lblCuentaDef.Text = & $textoCuenta ([int]$c.sldCuentaDef.Value); & $marcarPendiente })

    $c.btnGuardar.Add_Click({
            $c.btnGuardar.IsEnabled = $false
            $ventana.Cursor = [System.Windows.Input.Cursors]::Wait
            try {
                $valores = @{
                    Empresa            = $c.txtEmpresa.Text
                    ColorAcento        = $c.txtColor.Text
                    EstiloDialogo      = if ($c.cmbEstilo.SelectedIndex -eq 1) { 'Classic' } else { 'Fluent' }
                    Idioma             = switch ($c.cmbIdioma.SelectedIndex) { 0 { 'es' } 1 { 'en' } default { 'auto' } }
                    RutaLogs           = $c.txtLogs.Text
                    VecesAplazar       = [int]$c.sldAplazarDef.Value
                    DiasLimite         = [int]$c.sldDiasDef.Value
                    CuentaRegresivaMin = [int]$c.sldCuentaDef.Value
                    MensajeBienvenida  = $c.txtBienvenidaDef.Text
                    MensajeFinal       = $c.txtMensajeDef.Text
                    TextoBotonInstalar = $c.txtBotonInstalarDef.Text
                    TextoBotonAplazar  = $c.txtBotonAplazarDef.Text
                }
                Save-PlantillaCorporativa -Valores $valores -LogoPendiente $script:LogoPendiente -BannerPendiente $script:BannerPendiente | Out-Null
                & $cargarPlantilla
                [System.Windows.MessageBox]::Show("Plantilla guardada.`n`nSe aplicará a los paquetes que crees desde ahora. Los paquetes que ya existían no cambian: vuelve a crearlos.", 'Plantilla corporativa', 'OK', 'Information') | Out-Null
            }
            catch {
                Write-Registro -Mensaje $_.Exception.Message -Tipo Error
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'No se pudo guardar la plantilla', 'OK', 'Error') | Out-Null
            }
            finally {
                $c.btnGuardar.IsEnabled = $true
                $ventana.Cursor = $null
            }
        })
    $c.btnAbrirRecursos.Add_Click({
            if (Test-Path -LiteralPath $script:Recursos) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:Recursos) }
        })

    # ---------- Atajos de teclado ----------
    $ventana.Add_PreviewKeyDown({
            param($s, $e)
            $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -eq [System.Windows.Input.ModifierKeys]::Control
            $clic = New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
            if ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::O) {
                & $irA 'Crear'
                if ($c.zonaArrastre.IsEnabled) { & $elegirInstalador }
                $e.Handled = $true
            }
            elseif ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::Return -and $script:PaginaActual -eq 'Crear') {
                if ($c.btnCrear.IsEnabled) { $c.btnCrear.RaiseEvent($clic) }
                $e.Handled = $true
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::F5 -and $script:PaginaActual -eq 'Crear') {
                if ($c.btnVista.IsEnabled) { $c.btnVista.RaiseEvent($clic) }
                $e.Handled = $true
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::F5 -and $script:PaginaActual -eq 'Paquetes') {
                & $cargarPaquetes
                $e.Handled = $true
            }
        })

    # ---------- Inicio ----------
    $ventana.Add_ContentRendered({
            & $irA 'Crear'
            & $aplicarModo ''
            & $mostrarLogoApp '' ''
            $c.lblAplazar.Text = & $textoAplazar ([int]$c.sldAplazar.Value)
            $c.lblDias.Text = & $textoDias ([int]$c.sldDias.Value)
            $c.lblCuenta.Text = & $textoCuenta ([int]$c.sldCuenta.Value)
            try {
                Initialize-Fabrica
                $c.txtVersionPsadt.Text = "PSADT $($script:VersionPSADT)"
                try { & $cargarPlantilla }
                catch {
                    $script:Listo = $true
                    Write-Registro -Mensaje "Aviso al mostrar la plantilla (la Fábrica sigue funcionando): $($_.Exception.Message)" -Tipo Aviso
                }
                & $irA 'Crear'
                Write-Registro -Mensaje 'Listo. Arrastra un instalador para empezar.' -Tipo OK
            }
            catch {
                $c.expRegistro.IsExpanded = $true
                Write-Registro -Mensaje $_.Exception.Message -Tipo Error
                foreach ($b in @($c.btnCrear, $c.btnVista, $c.btnGuardar)) { $b.IsEnabled = $false }
                $c.zonaArrastre.IsEnabled = $false
                [System.Windows.MessageBox]::Show($_.Exception.Message, 'No se pudo preparar la Fábrica', 'OK', 'Error') | Out-Null
            }
        })

    [void]$ventana.ShowDialog()
    $script:CajaRegistro = $null
    $script:TextoEstado = $null
}
#endregion

#region ===================== Modo web =====================
# Servidor HTTP mínimo sobre TcpListener, SOLO en 127.0.0.1 (no acepta conexiones de otros equipos).
# No usa HttpListener (que pide reservar la URL con permisos de administrador) ni IIS.
# Seguridad: cabecera Host y Origin validadas (evita DNS rebinding y peticiones de otras páginas) y
# token aleatorio por sesión en todas las peticiones /api. La página se sirve con CSP estricta.
$script:HtmlWeb = @'
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fábrica de Paquetes</title>
<style>
:root{
  --acento:#2C2A29; --acento-texto:#fff; --acento-suave:rgba(44,42,41,.09);
  --fondo:#F4F6F9; --tarjeta:#fff; --borde:#E4E7EC; --borde-campo:#D0D5DD;
  --texto:#101828; --texto2:#344054; --texto3:#667085; --texto4:#98A2B3; --campo:#fff; --campo-ro:#F2F4F7;
  --ok:#027A48; --ok-fondo:#ECFDF3; --aviso:#B54708; --aviso-fondo:#FFFAEB; --error:#B42318; --error-fondo:#FEF3F2;
  --sombra:0 1px 2px rgba(16,24,40,.06),0 1px 3px rgba(16,24,40,.08);
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --fondo:#0F1115; --tarjeta:#181B21; --borde:#2A2F38; --borde-campo:#3A404B;
    --texto:#F2F4F7; --texto2:#D0D5DD; --texto3:#98A2B3; --texto4:#667085; --campo:#111318; --campo-ro:#1F232A;
    --ok:#6CE9A6; --ok-fondo:#0B2A1C; --aviso:#FEC84B; --aviso-fondo:#2B2111; --error:#FDA29B; --error-fondo:#2D1414;
    --sombra:0 1px 3px rgba(0,0,0,.4);
  }
}
:root[data-theme="dark"]{
  --fondo:#0F1115; --tarjeta:#181B21; --borde:#2A2F38; --borde-campo:#3A404B;
  --texto:#F2F4F7; --texto2:#D0D5DD; --texto3:#98A2B3; --texto4:#667085; --campo:#111318; --campo-ro:#1F232A;
  --ok:#6CE9A6; --ok-fondo:#0B2A1C; --aviso:#FEC84B; --aviso-fondo:#2B2111; --error:#FDA29B; --error-fondo:#2D1414;
  --sombra:0 1px 3px rgba(0,0,0,.4);
}
*{box-sizing:border-box}
html,body{margin:0;height:100%}
body{background:var(--fondo);color:var(--texto);font:14px/1.45 "Segoe UI",system-ui,-apple-system,sans-serif;display:flex;flex-direction:column}
button,input,select,textarea{font:inherit;color:inherit}
header{background:var(--acento);color:var(--acento-texto);padding:12px 22px;display:flex;align-items:center;gap:14px;flex:none}
header .logo{width:40px;height:40px;border-radius:8px;background:#fff;display:grid;place-items:center;overflow:hidden;flex:none}
header .logo img{max-width:32px;max-height:32px}
header h1{font-size:19px;font-weight:600;margin:0;line-height:1.2}
header .sub{font-size:12px;opacity:.85}
header .der{margin-left:auto;display:flex;gap:8px;align-items:center}
.pastilla{background:rgba(255,255,255,.2);border-radius:12px;padding:4px 12px;font-size:12px;white-space:nowrap}
.btn-cab{white-space:nowrap;background:rgba(255,255,255,.18);border:0;color:inherit;border-radius:8px;padding:6px 12px;cursor:pointer;font-size:12.5px}
.btn-cab:hover{background:rgba(255,255,255,.3)}
.cuerpo{flex:1;display:flex;min-height:0}
nav{width:216px;flex:none;background:var(--tarjeta);border-right:1px solid var(--borde);padding:16px 12px;overflow:auto}
nav .item{display:flex;gap:10px;align-items:center;width:100%;text-align:left;border:0;background:transparent;padding:10px 14px;border-radius:8px;cursor:pointer;color:var(--texto2);margin-bottom:4px;font-size:14px}
nav .item:hover{background:var(--acento-suave)}
nav .item.activo{background:var(--acento-suave);font-weight:600;color:var(--texto)}
nav hr{border:0;border-top:1px solid var(--borde);margin:14px 0}
.titulo-nav{font-size:11px;font-weight:600;color:var(--texto4);margin:0 6px 6px;letter-spacing:.04em}
.paso{font-size:12.5px;color:var(--texto4);padding:3px 6px}
.paso.ok{color:var(--ok);font-weight:600}
.atajos{font-size:11.5px;color:var(--texto3);margin:12px 6px 0;white-space:pre-line}
main{flex:1;overflow:auto;padding:16px 18px 8px;min-width:0}
.pagina{display:none}
.pagina.activa{display:block}
.dos{display:grid;grid-template-columns:minmax(0,1fr) 400px;gap:18px;align-items:start}
.panel-der{position:sticky;top:0;max-height:calc(100vh - 140px);overflow:auto;padding-right:4px}
.tarjeta{background:var(--tarjeta);border:1px solid var(--borde);border-radius:10px;padding:16px 20px;margin-bottom:14px;box-shadow:var(--sombra)}
.tarjeta h2{font-size:15px;font-weight:600;margin:0 0 2px}
.tarjeta .subt{font-size:12px;color:var(--texto3);margin:0 0 12px}
label.et{display:block;font-weight:600;color:var(--texto2);margin:10px 0 4px}
.ayuda{font-size:11.5px;color:var(--texto3);margin-top:3px}
.msg{font-size:11.5px;margin-top:4px;display:none}
.msg.error{display:block;color:var(--error)}
.msg.aviso{display:block;color:var(--aviso)}
input[type=text],textarea,select{width:100%;background:var(--campo);border:1px solid var(--borde-campo);border-radius:6px;padding:7px 9px;outline:none}
textarea{resize:vertical;min-height:56px}
input[type=text]:focus,textarea:focus,select:focus{border-color:var(--acento);box-shadow:0 0 0 3px var(--acento-suave)}
input.error,textarea.error{border-color:#D92D20;border-width:1.5px}
input.aviso,textarea.aviso{border-color:#F79009;border-width:1.5px}
input[readonly]{background:var(--campo-ro)}
.fila{display:grid;grid-template-columns:2fr 1fr;gap:14px}
.fila.mitad{grid-template-columns:1fr 1fr}
.con-boton{display:flex;gap:8px}
.chk{display:flex;gap:8px;align-items:flex-start;margin-top:10px;color:var(--texto2);cursor:pointer}
.chk input{margin-top:3px;accent-color:var(--acento)}
.chk.deshabilitado{opacity:.5;cursor:default}
.btn{border:1px solid var(--borde-campo);background:var(--tarjeta);color:var(--texto2);font-weight:600;padding:7px 14px;border-radius:8px;cursor:pointer;white-space:nowrap}
.btn:hover{background:var(--campo-ro)}
.btn:disabled{opacity:.45;cursor:default}
.btn.primario{background:var(--acento);color:var(--acento-texto);border-color:transparent;font-size:15px;padding:11px 18px;width:100%}
.btn.primario:hover{filter:brightness(1.08)}
.btn.ancho{width:100%}
.zona{border:1.5px dashed var(--acento);background:var(--acento-suave);border-radius:10px;padding:18px;text-align:center;cursor:pointer;min-height:96px;display:flex;flex-direction:column;justify-content:center;transition:background .15s}
.zona.encima{background:var(--acento);color:var(--acento-texto)}
.zona .ico{font-size:26px;line-height:1}
.zona .t1{font-weight:600;font-size:14px}
.zona .t2{font-size:12px;color:var(--texto3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.zona.encima .t2{color:inherit}
.chip{display:inline-block;background:var(--campo-ro);border-radius:10px;padding:2px 10px;font-weight:600;font-size:12px;color:var(--texto2)}
.logo-caja{width:60px;height:60px;border-radius:10px;border:1px solid var(--borde);background:var(--campo-ro);display:grid;place-items:center;overflow:hidden;flex:none;font-weight:700;font-size:20px;color:var(--texto4)}
.logo-caja img{max-width:48px;max-height:48px}
.deslizadores{display:grid;grid-template-columns:190px 1fr 100px;gap:6px 12px;align-items:center;margin-top:14px}
.deslizadores span.et2{font-weight:600;color:var(--texto2)}
.deslizadores b{font-weight:600}
input[type=range]{width:100%;accent-color:var(--acento)}
.oculto{display:none!important}
/* Maqueta de la ventana de PSADT */
.maqueta{background:#fff;color:#101828;border:1px solid #D0D5DD;border-radius:8px;display:grid;grid-template-columns:6px 1fr;box-shadow:0 6px 18px rgba(0,0,0,.18);overflow:hidden}
.maqueta .franja{background:var(--acento)}
.maqueta .dentro{padding:14px}
.maqueta .cab{display:flex;gap:10px;align-items:center}
.maqueta .ini{width:38px;height:38px;border-radius:8px;background:var(--acento);color:var(--acento-texto);display:grid;place-items:center;font-weight:700;overflow:hidden;flex:none}
.maqueta .ini img{width:38px;height:38px;object-fit:contain;background:#fff}
.maqueta .tit{font-weight:600;font-size:15px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.maqueta .st{font-size:11px;color:#667085}
.maqueta .caja{border:1px solid #E4E7EC;border-radius:6px;margin-top:12px;padding:10px;font-size:11.5px;color:#344054}
#mockTexto{white-space:pre-line}
.maqueta .lista{border:1px solid #E4E7EC;border-radius:6px;margin-top:10px;padding:4px 8px}
.maqueta .proc{display:flex;gap:8px;align-items:center;margin:4px 0}
.maqueta .proc i{width:22px;height:22px;border-radius:4px;background:var(--acento-suave);display:grid;place-items:center;font-style:normal;font-weight:700;font-size:11px}
.maqueta .proc b{display:block;font-size:11.5px}
.maqueta .proc small{font-size:10px;color:var(--acento)}
.maqueta .fila-m{display:flex;justify-content:space-between;font-size:11.5px;color:#667085;margin-top:6px}
.maqueta .fila-m b{color:#101828}
.maqueta .fila-m b.res{color:#B8860B}
.maqueta .botones{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:12px}
.maqueta .botones div{border-radius:4px;padding:6px;text-align:center;font-size:11.5px}
.maqueta .bi{border:1px solid #D0D5DD;color:#344054}
.maqueta .ba{background:var(--acento);color:var(--acento-texto)}
.maqueta .acento,.maqueta .url{color:var(--acento)}
.maqueta .url{text-decoration:underline}
.nota{border-radius:8px;padding:9px 12px;font-size:12px;margin-top:12px}
.nota.naranja{background:var(--aviso-fondo);color:var(--aviso)}
.nota.gris{background:var(--campo-ro);color:var(--texto2)}
.nota.verde{background:var(--ok-fondo);color:var(--ok)}
.nota.roja{background:var(--error-fondo);color:var(--error)}
.caja-destino{background:var(--tarjeta);border:1px solid var(--borde);border-radius:8px;padding:10px 12px;margin:14px 0 10px}
.caja-destino .ruta{font-size:11.5px;color:var(--texto3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin:4px 0 8px}
.dos-botones{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px}
/* Paquetes */
.barra-pag{display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap}
.barra-pag h2{font-size:18px;margin:0}
.barra-pag .espacio{flex:1}
.barra-pag input{width:260px}
.paquete{display:grid;grid-template-columns:44px minmax(0,1fr) auto;gap:14px;align-items:center}
.paquete .tipo{width:44px;height:44px;border-radius:8px;background:var(--acento-suave);display:grid;place-items:center;font-weight:700;font-size:11px}
.paquete .n{font-weight:600;font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.paquete .d{font-size:11.5px;color:var(--texto3)}
.paquete .c{font-size:11px;color:var(--texto4);font-family:Consolas,monospace}
.paquete .acc{display:flex;gap:8px}
/* Plantilla */
.colores{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.colores input[type=color]{width:38px;height:34px;padding:0;border:1px solid var(--borde-campo);border-radius:6px;background:none}
.colores input[type=text]{width:110px}
.muestra{width:24px;height:24px;border-radius:5px;border:1px solid var(--borde-campo);cursor:pointer}
.banner-caja{border:1px solid var(--borde);background:var(--campo-ro);border-radius:6px;padding:4px;display:inline-block;min-width:200px;min-height:58px}
.banner-caja img{max-width:360px;max-height:50px;display:block}
.prev-icono{width:110px;height:110px;border-radius:10px;border:1px solid var(--borde);background:var(--campo-ro);display:grid;place-items:center}
.prev-icono img{max-width:94px;max-height:94px}
.ayuda-pag p{margin:4px 0;color:var(--texto2)}
/* Barra inferior */
footer{flex:none;background:var(--tarjeta);border-top:1px solid var(--borde);padding:8px 18px 8px 234px}
footer .estado{cursor:pointer;color:var(--texto3);display:flex;gap:8px;align-items:center;user-select:none}
footer .estado .flecha{font-size:10px;transition:transform .15s}
footer.abierto .estado .flecha{transform:rotate(90deg)}
footer pre{display:none;height:150px;overflow:auto;margin:8px 0 0;background:var(--campo);border:1px solid var(--borde-campo);border-radius:6px;padding:8px;font:11.5px/1.5 Consolas,monospace;white-space:pre}
footer.abierto pre{display:block}
.l-OK{color:var(--ok)} .l-Aviso{color:var(--aviso)} .l-Error{color:var(--error)}
/* Capa de espera y diálogo */
.capa{position:fixed;inset:0;background:rgba(16,24,40,.45);display:none;align-items:center;justify-content:center;z-index:10;padding:16px}
.capa.visible{display:flex}
.dialogo{background:var(--tarjeta);border-radius:12px;padding:22px 24px;max-width:520px;width:100%;box-shadow:0 20px 40px rgba(0,0,0,.3)}
.dialogo h3{margin:0 0 8px;font-size:17px}
.dialogo .txt{white-space:pre-line;color:var(--texto2);font-size:13.5px;word-break:break-word}
.dialogo .acc{display:flex;gap:8px;justify-content:flex-end;margin-top:18px}
.dialogo .acc .btn.primario{width:auto;font-size:14px;padding:8px 18px}
.giro{width:34px;height:34px;border-radius:50%;border:4px solid var(--acento-suave);border-top-color:var(--acento);animation:g 0.9s linear infinite;margin:0 auto 12px}
@keyframes g{to{transform:rotate(360deg)}}
.centro{text-align:center}
@media (max-width:1100px){.dos{grid-template-columns:1fr}.panel-der{position:static;max-height:none;overflow:visible}}
@media (max-width:760px){
  nav{display:none} footer{padding-left:16px} main{padding:12px 16px}
  .fila,.fila.mitad{grid-template-columns:1fr} .deslizadores{grid-template-columns:1fr 70px} .deslizadores span.et2{grid-column:1/-1}
  header{padding:10px 16px} header .pastilla,header .sub{display:none} .paquete{grid-template-columns:44px minmax(0,1fr)} .paquete .acc{grid-column:1/-1}
  .barra-pag input{width:100%}
}
</style>
</head>
<body>
<header>
  <div class="logo"><img id="imgCabecera" alt=""></div>
  <div><h1>Fábrica de Paquetes</h1><div class="sub">PSAppDeployToolkit → Intune · versión web</div></div>
  <div class="der">
    <span class="pastilla" id="txtVersionPsadt">PSADT: cargando…</span>
    <button class="btn-cab" id="btnTema" title="Cambiar entre tema claro y oscuro">◐ Tema</button>
    <button class="btn-cab" id="btnSalir" title="Detener la Fábrica web y cerrar la consola">⏻ Cerrar Fábrica</button>
  </div>
</header>
<div class="cuerpo">
  <nav>
    <button class="item activo" data-pag="crear">📦 Crear paquete</button>
    <button class="item" data-pag="paquetes">🗂 Paquetes creados</button>
    <button class="item" data-pag="plantilla">🎨 Plantilla corporativa</button>
    <button class="item" data-pag="ayuda">❓ Cómo funciona</button>
    <hr>
    <div class="titulo-nav">TU AVANCE</div>
    <div class="paso" id="paso1">○ 1. Instalador</div>
    <div class="paso" id="paso2">○ 2. Información de la app</div>
    <div class="paso" id="paso3">○ 3. Cómo se instala</div>
    <div class="paso" id="paso4">○ 4. Ventanas del usuario</div>
    <div class="atajos">Atajos:
Ctrl+O  elegir instalador
Ctrl+Enter  crear paquete</div>
  </nav>
  <main>
    <!-- ===================== Crear paquete ===================== -->
    <section class="pagina activa" id="pag-crear">
      <div class="dos">
        <div>
          <div class="tarjeta">
            <h2>1 · Instalador</h2>
            <p class="subt">Arrastra aquí el instalador o haz clic para buscarlo en el equipo. Se leen solos el nombre, el fabricante y la versión.</p>
            <div class="zona" id="zona" tabindex="0">
              <div class="ico" id="zonaIco">⬇</div>
              <div class="t1" id="zonaT1">Arrastra aquí el .msi, .exe o .ps1</div>
              <div class="t2" id="zonaT2">o haz clic para buscarlo</div>
            </div>
            <label class="et" for="txtRuta">O pega la ruta completa del instalador</label>
            <div class="con-boton">
              <input type="text" id="txtRuta" placeholder="D:\Fuentes\MiApp\setup.msi">
              <button class="btn" id="btnUsarRuta">Leer</button>
            </div>
            <div class="ayuda">Usa la ruta (o «clic para buscar») si el instalador trae archivos adicionales o es un .ps1 con su script de desinstalación: al arrastrar solo se copia ese archivo.</div>
            <div style="margin-top:10px"><span class="chip oculto" id="chipTipo"></span> <span class="ayuda" id="txtTipoDetalle"></span></div>
            <label class="chk" id="lblCarpeta"><input type="checkbox" id="chkCarpeta"> Copiar toda la carpeta del instalador (si trae archivos adicionales)</label>
          </div>

          <div class="tarjeta">
            <h2>2 · Información de la app</h2>
            <p class="subt">Así aparece la aplicación en Intune y en Company Portal.</p>
            <div class="fila">
              <div><label class="et" for="txtNombre">Nombre de la aplicación</label><input type="text" id="txtNombre"><div class="msg" id="msgNombre"></div></div>
              <div><label class="et" for="txtVersion">Versión</label><input type="text" id="txtVersion"><div class="msg" id="msgVersion"></div></div>
            </div>
            <div class="fila">
              <div><label class="et" for="txtFabricante">Fabricante (publicador)</label><input type="text" id="txtFabricante"></div>
              <div><label class="et" for="cmbArq">Arquitectura</label><select id="cmbArq"><option>x64</option><option>x86</option></select></div>
            </div>
            <label class="et" for="txtDescripcion">Descripción para Company Portal</label>
            <textarea id="txtDescripcion" rows="3"></textarea>
            <div class="ayuda">Qué hace la aplicación y para quién es. Intune la exige; si la dejas vacía se escribe una automática.</div>
            <label class="et">Logo de la app</label>
            <div style="display:flex;gap:14px;align-items:center">
              <div class="logo-caja" id="logoAppCaja"><span id="logoAppIni">AP</span></div>
              <div style="flex:1">
                <div class="ayuda" id="txtLogoAppInfo" style="margin:0 0 6px">Sin logo: se generan las iniciales con el color de la empresa.</div>
                <button class="btn" id="btnLogoApp">Elegir imagen…</button>
                <button class="btn" id="btnQuitarLogoApp">Quitar</button>
                <label class="chk" id="lblLogoVentanas"><input type="checkbox" id="chkLogoVentanas"> Mostrar este logo en las ventanas (en vez del logo de la empresa)</label>
              </div>
            </div>
          </div>

          <div class="tarjeta">
            <h2>3 · Cómo se instala</h2>
            <p class="subt" id="txtSubInstalacion">Primero elige un instalador.</p>
            <div id="filaArgs" class="oculto">
              <label class="et" for="txtArgs" id="lblArgs">Argumentos de instalación</label>
              <input type="text" id="txtArgs"><div class="msg" id="msgArgs"></div>
              <div class="ayuda" id="lblAyudaArgs"></div>
            </div>
            <div id="filaDes" class="oculto">
              <label class="et" for="txtDes" id="lblDes">Desinstalador</label>
              <div class="con-boton"><input type="text" id="txtDes"><button class="btn" id="btnDes">Buscar…</button></div>
              <div class="msg" id="msgDes"></div>
              <div class="ayuda" id="lblAyudaDes"></div>
            </div>
            <div id="filaArgsDes" class="oculto">
              <label class="et" for="txtArgsDes">Argumentos de desinstalación</label>
              <input type="text" id="txtArgsDes">
            </div>
            <div id="filaDet" class="oculto">
              <label class="et" for="txtDeteccion">Nombre para la detección</label>
              <input type="text" id="txtDeteccion">
              <div class="ayuda">Como aparece en Configuración → Aplicaciones (Programas y características).</div>
            </div>
          </div>

          <div class="tarjeta">
            <h2>4 · Ventanas que ve el usuario</h2>
            <p class="subt">Mira cómo cambia la vista previa mientras ajustas.</p>
            <label class="et" for="txtProcesos">Aplicaciones que se deben cerrar</label>
            <input type="text" id="txtProcesos">
            <div class="ayuda">Separadas por coma, sin .exe. Puedes darles nombre visible: excel=Microsoft Excel, AppCaja=Aplicativo Caja</div>
            <label class="chk"><input type="checkbox" id="chkSiempre"> Mostrar las ventanas aunque la app esté cerrada (si no, se instala en silencio)</label>
            <div class="deslizadores">
              <span class="et2">Veces que puede aplazar</span><input type="range" id="sldAplazar" min="0" max="10" value="3"><b id="lblAplazar"></b>
              <span class="et2">Fecha límite para aplazar</span><input type="range" id="sldDias" min="0" max="30" value="0"><b id="lblDias"></b>
              <span class="et2">Cuenta regresiva</span><input type="range" id="sldCuenta" min="0" max="60" value="0"><b id="lblCuenta"></b>
            </div>
            <div class="ayuda">Cuenta regresiva: al llegar a cero se cierran las apps y se instala solo.</div>
            <label class="et" for="txtBienvenida">Texto de la ventana (opcional)</label>
            <textarea id="txtBienvenida" rows="2"></textarea>
            <div class="ayuda">Si escribes algo, REEMPLAZA el texto estándar de PSADT. Pulsa Enter para un salto de línea. Formato: [bold]negrita[/bold]  [italic]cursiva[/italic]  [accent]color[/accent]  [url]https://…[/url]</div>
            <div class="fila mitad">
              <div><label class="et" for="txtBotonInstalar">Texto del botón Instalar</label><input type="text" id="txtBotonInstalar"></div>
              <div><label class="et" for="txtBotonAplazar">Texto del botón Aplazar</label><input type="text" id="txtBotonAplazar"></div>
            </div>
            <div class="ayuda">Vacío = texto estándar de PSADT. Usa textos cortos: el botón es angosto.</div>
            <label class="et" for="txtMensaje">Mensaje al terminar (opcional)</label>
            <input type="text" id="txtMensaje">
          </div>
        </div>

        <div class="panel-der">
          <div style="font-size:15px;font-weight:600">Así lo verá el usuario</div>
          <div class="ayuda" style="margin-bottom:10px">Vista aproximada. Para ver las ventanas reales de PSADT usa «Ver ventanas reales» (no instala nada).</div>
          <div class="maqueta">
            <div class="franja"></div>
            <div class="dentro">
              <div class="cab">
                <div class="ini" id="mockIni">AP</div>
                <div style="min-width:0"><div class="tit" id="mockTitulo">Nombre de la app 1.0</div><div class="st" id="mockSub">Empresa - Instalación de la aplicación</div></div>
              </div>
              <div class="caja">
                <div id="mockTexto"></div>
                <div class="lista" id="mockLista"></div>
              </div>
              <div id="mockFilas"></div>
              <div class="botones" id="mockBotones"><div class="bi" id="mockBtnInstalar">Instalar</div><div class="ba" id="mockBtnAplazar">Aplazar</div></div>
            </div>
          </div>
          <div class="nota naranja" id="txtAvisoSilencio"></div>

          <div class="caja-destino">
            <div style="font-weight:600;color:var(--texto2)">📁 Dónde queda el paquete listo</div>
            <div class="ruta" id="txtDestino"></div>
            <button class="btn" id="btnDestino">Cambiar carpeta…</button>
            <button class="btn" id="btnDestinoDefecto">Usar la de la Fábrica</button>
            <label class="chk"><input type="checkbox" id="chkCopiarFuente"> Incluir también la carpeta fuente (para editarlo después)</label>
          </div>
          <label class="chk" style="margin:0 0 10px"><input type="checkbox" id="chkIntunewin" checked> Generar también el archivo .intunewin</label>
          <div class="nota gris" id="txtResumen" style="margin:0 0 10px">Arrastra un instalador para empezar.</div>
          <button class="btn primario" id="btnCrear">Crear paquete</button>
          <div class="dos-botones">
            <button class="btn" id="btnVista">👁 Ver ventanas reales</button>
            <button class="btn" id="btnAbrir" disabled>📂 Abrir salida</button>
          </div>
        </div>
      </div>
    </section>

    <!-- ===================== Paquetes creados ===================== -->
    <section class="pagina" id="pag-paquetes">
      <div class="barra-pag">
        <div><h2>Paquetes creados</h2><div class="ayuda" id="txtPaquetesResumen">Del más reciente al más antiguo.</div></div>
        <div class="espacio"></div>
        <input type="text" id="txtBuscarPaquete" placeholder="Buscar por nombre, versión o fabricante">
        <button class="btn" id="btnRefrescarPaquetes">⟳ Actualizar</button>
      </div>
      <div id="listaPaquetes"></div>
    </section>

    <!-- ===================== Plantilla corporativa ===================== -->
    <section class="pagina" id="pag-plantilla">
      <div class="dos">
        <div>
          <div class="tarjeta">
            <h2>Identidad de la empresa</h2>
            <p class="subt">Se aplica a todos los paquetes que crees DESPUÉS de guardar.</p>
            <label class="et" for="txtEmpresa">Nombre de la empresa</label>
            <input type="text" id="txtEmpresa">
            <div class="ayuda">Sale en el subtítulo de las ventanas y en el banner.</div>
            <label class="et" for="txtColor">Color corporativo</label>
            <div class="colores">
              <input type="color" id="selColor" title="Elegir color">
              <input type="text" id="txtColor">
              <span id="paleta" class="colores"></span>
            </div>
            <div class="msg" id="msgColor"></div>
            <label class="et">Logo de la empresa</label>
            <div style="display:flex;gap:14px;align-items:center">
              <div class="logo-caja" id="logoEmpCaja"></div>
              <div>
                <div class="ayuda" id="txtLogoEmpInfo" style="margin:0 0 6px"></div>
                <button class="btn" id="btnLogo">Elegir imagen…</button>
                <button class="btn" id="btnQuitarLogo">Usar iniciales</button>
              </div>
            </div>
            <div class="ayuda">Ideal: PNG cuadrado con fondo transparente, 256 px o más. También puedes arrastrar la imagen sobre el cuadro.</div>
            <label class="et">Banner (solo se ve con el estilo Classic)</label>
            <div class="banner-caja" id="bannerEmpCaja"></div>
            <div class="ayuda" id="txtBannerEmpInfo" style="margin:6px 0"></div>
            <button class="btn" id="btnBanner">Subir banner…</button>
            <button class="btn" id="btnQuitarBanner">Generar automático</button>
            <div class="ayuda">Recomendado: PNG de 900 × 125 px (proporción 7,2 a 1). Con otra proporción se centra sin deformarse.</div>
          </div>
          <div class="tarjeta">
            <h2>Ventanas</h2>
            <div class="fila mitad">
              <div><label class="et" for="cmbEstilo">Estilo</label><select id="cmbEstilo"><option value="Fluent">Fluent</option><option value="Classic">Classic</option></select><div class="ayuda">Fluent = moderno (recomendado).</div></div>
              <div><label class="et" for="cmbIdioma">Idioma</label><select id="cmbIdioma"><option value="es">Español</option><option value="en">Inglés</option><option value="auto">Automático (idioma del equipo)</option></select></div>
            </div>
            <label class="et" for="txtLogs">Ruta de logs en el equipo</label>
            <input type="text" id="txtLogs">
            <div class="ayuda">Por defecto $envWinDir\Logs\Software  (C:\Windows\Logs\Software).</div>
          </div>
          <div class="tarjeta">
            <h2>Valores por defecto de cada paquete nuevo</h2>
            <div class="deslizadores">
              <span class="et2">Veces que puede aplazar</span><input type="range" id="sldAplazarDef" min="0" max="10"><b id="lblAplazarDef"></b>
              <span class="et2">Fecha límite para aplazar</span><input type="range" id="sldDiasDef" min="0" max="30"><b id="lblDiasDef"></b>
              <span class="et2">Cuenta regresiva</span><input type="range" id="sldCuentaDef" min="0" max="60"><b id="lblCuentaDef"></b>
            </div>
            <label class="et" for="txtBienvenidaDef">Texto de la ventana por defecto (reemplaza el estándar de PSADT)</label>
            <textarea id="txtBienvenidaDef" rows="2"></textarea>
            <div class="fila mitad">
              <div><label class="et" for="txtBotonInstalarDef">Botón Instalar por defecto</label><input type="text" id="txtBotonInstalarDef"></div>
              <div><label class="et" for="txtBotonAplazarDef">Botón Aplazar por defecto</label><input type="text" id="txtBotonAplazarDef"></div>
            </div>
            <label class="et" for="txtMensajeDef">Mensaje final por defecto</label>
            <input type="text" id="txtMensajeDef">
          </div>
        </div>
        <div class="panel-der">
          <div style="font-size:15px;font-weight:600">Así quedan los recursos</div>
          <div class="ayuda" style="margin-bottom:12px">Se regeneran al pulsar «Guardar plantilla».</div>
          <div class="tarjeta">
            <label class="et" style="margin-top:0">Logo de las ventanas (Fluent)</label>
            <div class="prev-icono"><img id="imgIconoPrev" alt=""></div>
            <label class="et" style="margin-top:14px">Banner (estilo Classic)</label>
            <img id="imgBannerPrev" alt="" style="max-width:100%;max-height:60px">
          </div>
          <div class="nota naranja oculto" id="bdPendiente" style="margin:0 0 12px">Tienes cambios sin guardar. Pulsa «Guardar plantilla».</div>
          <button class="btn primario" id="btnGuardar">Guardar plantilla</button>
          <button class="btn ancho" id="btnAbrirRecursos" style="margin-top:8px">📂 Abrir carpeta Recursos</button>
        </div>
      </div>
    </section>

    <!-- ===================== Cómo funciona ===================== -->
    <section class="pagina ayuda-pag" id="pag-ayuda" style="max-width:860px">
      <div class="tarjeta"><h2>Crear un paquete en 4 pasos</h2>
        <p>1. Arrastra el instalador (.msi, .exe o .ps1), haz clic para buscarlo o pega su ruta.</p>
        <p>2. Revisa nombre, versión, fabricante y escribe la descripción para Company Portal.</p>
        <p>3. Ajusta las ventanas (apps a cerrar, aplazamientos, cuenta regresiva) y mira la vista previa.</p>
        <p>4. Pulsa «Crear paquete». En la carpeta de salida quedan el .intunewin, FichaIntune.txt, Deteccion.ps1, Descripcion.txt y Logo_Intune.png.</p></div>
      <div class="tarjeta"><h2>¿Por qué a veces no salen las ventanas?</h2>
        <p>PSADT decide solo (modo Auto):</p>
        <p>• Si hay apps a cerrar y están ABIERTAS → muestra la ventana completa.</p>
        <p>• Si hay apps a cerrar y están CERRADAS → instala en silencio (no molesta al usuario).</p>
        <p>• Si no hay nadie conectado → siempre en silencio.</p>
        <p>Para que salgan siempre, marca «Mostrar las ventanas aunque la app esté cerrada».</p></div>
      <div class="tarjeta"><h2>Tipos de instalador</h2>
        <p>• MSI: todo es automático. Se detecta por su código de producto.</p>
        <p>• EXE: escribe los argumentos silenciosos (/S, /quiet…) y la ruta del desinstalador en el equipo.</p>
        <p>• PS1: se copia toda la carpeta del script; debe terminar con «exit 0». Se detecta por una marca en el registro que crea el paquete.</p></div>
      <div class="tarjeta"><h2>Sobre la versión web</h2>
        <p>• Funciona solo en este equipo: el servidor escucha en 127.0.0.1 y pide un código secreto que cambia cada vez que se abre.</p>
        <p>• La ventana negra (consola) es la Fábrica: no la cierres mientras uses la página. Para terminar pulsa «Cerrar Fábrica».</p>
        <p>• Mientras se crea un paquete la página espera; no hace falta recargarla.</p></div>
      <div class="tarjeta"><h2>Antes de publicar en Intune</h2>
        <p>Prueba como SYSTEM con los comandos de FichaIntune.txt (psexec -s -i cmd). Instala, revisa que Deteccion.ps1 diga «Detectado», desinstala y vuelve a revisar. Luego publica en un grupo piloto.</p></div>
    </section>
  </main>
</div>
<footer id="pie">
  <div class="estado" id="estado"><span class="flecha">▶</span><span id="txtEstado">Conectando con la Fábrica…</span></div>
  <pre id="txtLog"></pre>
</footer>

<div class="capa" id="capaEspera"><div class="dialogo centro"><div class="giro"></div><h3 id="esperaTitulo">Trabajando…</h3><div class="txt" id="esperaTexto">No cierres la consola de la Fábrica.</div></div></div>
<div class="capa" id="capaDialogo"><div class="dialogo"><h3 id="dlgTitulo"></h3><div class="txt" id="dlgTexto"></div><div class="acc" id="dlgAcciones"></div></div></div>

<script>
(function () {
  'use strict';
  var TOKEN = new URLSearchParams(location.search).get('t') || '';
  var $ = function (id) { return document.getElementById(id); };
  var estado = {
    listo: false, instalador: '', tipo: '', logoApp: null, destino: '', ultimaSalida: '',
    cfg: {}, salidaPorDefecto: '', logoPendiente: null, bannerPendiente: null, version: 0, ocupado: false
  };

  // ---------------- Comunicación con la Fábrica ----------------
  function agregarRegistro(lineas) {
    if (!lineas || !lineas.length) return;
    var pre = $('txtLog');
    lineas.forEach(function (l) {
      var s = document.createElement('span');
      s.className = 'l-' + l.Tipo;
      s.textContent = l.Linea + '\n';
      pre.appendChild(s);
    });
    pre.scrollTop = pre.scrollHeight;
    var ultima = lineas[lineas.length - 1];
    var e = $('txtEstado');
    e.textContent = ultima.Mensaje;
    e.className = 'l-' + ultima.Tipo;
  }
  function api(ruta, cuerpo) {
    var op = { method: cuerpo === undefined ? 'GET' : 'POST', headers: { 'X-Fabrica-Token': TOKEN } };
    if (cuerpo !== undefined) { op.headers['Content-Type'] = 'application/json'; op.body = JSON.stringify(cuerpo); }
    return fetch(ruta, op).then(function (r) {
      return r.json().catch(function () { return { ok: false, error: 'Respuesta no válida de la Fábrica (HTTP ' + r.status + ').' }; });
    }, function () {
      return { ok: false, desconectado: true, error: 'No hay conexión con la Fábrica. ¿Se cerró la consola? Vuelve a abrir Crear-Paquete-Web.cmd.' };
    }).then(function (d) {
      agregarRegistro(d.registro);
      if (d.desconectado) { $('txtEstado').textContent = d.error; $('txtEstado').className = 'l-Error'; }
      return d;
    });
  }
  function subirArchivo(archivo, clase) {
    var url = '/api/subir?clase=' + encodeURIComponent(clase) + '&nombre=' + encodeURIComponent(archivo.name);
    return fetch(url, { method: 'POST', headers: { 'X-Fabrica-Token': TOKEN, 'Content-Type': 'application/octet-stream' }, body: archivo })
      .then(function (r) { return r.json(); }, function () { return { ok: false, error: 'No hay conexión con la Fábrica.' }; })
      .then(function (d) { agregarRegistro(d.registro); return d; });
  }
  // Las imágenes por id no cambian; las de la plantilla (tipo=...) llevan un número para no usar una copia vieja.
  function urlImagen(params) {
    var u = '/api/imagen?t=' + encodeURIComponent(TOKEN) + '&' + params;
    return params.indexOf('tipo=') === 0 ? u + '&v=' + (++estado.version) : u;
  }

  // ---------------- Diálogos ----------------
  function esperar(titulo, texto) {
    $('esperaTitulo').textContent = titulo;
    $('esperaTexto').textContent = texto || 'No cierres la consola de la Fábrica.';
    $('capaEspera').classList.add('visible');
    estado.ocupado = true;
  }
  function finEspera() { $('capaEspera').classList.remove('visible'); estado.ocupado = false; }
  function dialogo(titulo, texto, botones) {
    return new Promise(function (resolver) {
      $('dlgTitulo').textContent = titulo;
      $('dlgTexto').textContent = texto;
      var acc = $('dlgAcciones');
      acc.innerHTML = '';
      (botones || [{ t: 'Aceptar', v: true, p: true }]).forEach(function (b) {
        var el = document.createElement('button');
        el.className = 'btn' + (b.p ? ' primario' : '');
        el.textContent = b.t;
        el.onclick = function () { $('capaDialogo').classList.remove('visible'); resolver(b.v); };
        acc.appendChild(el);
      });
      $('capaDialogo').classList.add('visible');
      var principal = acc.querySelector('.primario');
      if (principal) principal.focus();
    });
  }
  function error(titulo, d) { return dialogo(titulo, (d && d.error) || 'Error desconocido.'); }

  // ---------------- Navegación ----------------
  function irA(pag) {
    document.querySelectorAll('nav .item').forEach(function (b) { b.classList.toggle('activo', b.dataset.pag === pag); });
    document.querySelectorAll('.pagina').forEach(function (s) { s.classList.toggle('activa', s.id === 'pag-' + pag); });
    estado.pagina = pag;
    if (pag === 'paquetes') cargarPaquetes();
  }
  document.querySelectorAll('nav .item').forEach(function (b) { b.onclick = function () { irA(b.dataset.pag); }; });
  $('estado').onclick = function () { $('pie').classList.toggle('abierto'); };

  // ---------------- Utilidades ----------------
  function iniciales(texto) {
    var p = String(texto || '').split(/[\s_\-.]+/).filter(function (x) { return /^[A-Za-z0-9]/.test(x); });
    if (p.length >= 2) return (p[0][0] + p[1][0]).toUpperCase();
    if (p.length === 1) return p[0].substring(0, 2).toUpperCase();
    return 'AP';
  }
  function hexValido(h) { return /^#?[0-9A-Fa-f]{6}$/.test(String(h || '').trim()); }
  function colorTexto(hex) {
    var h = hex.replace('#', ''); var r = parseInt(h.substr(0, 2), 16), g = parseInt(h.substr(2, 2), 16), b = parseInt(h.substr(4, 2), 16);
    return ((0.299 * r + 0.587 * g + 0.114 * b) / 255) > 0.6 ? '#1E1E1E' : '#FFFFFF';
  }
  function aplicarAcento(hex) {
    if (!hexValido(hex)) return;
    hex = '#' + hex.trim().replace('#', '');
    var h = hex.replace('#', ''); var r = parseInt(h.substr(0, 2), 16), g = parseInt(h.substr(2, 2), 16), b = parseInt(h.substr(4, 2), 16);
    var st = document.documentElement.style;
    st.setProperty('--acento', hex);
    st.setProperty('--acento-texto', colorTexto(hex));
    st.setProperty('--acento-suave', 'rgba(' + r + ',' + g + ',' + b + ',.12)');
  }
  function listaProcesos(texto, descripcion) {
    var partes = String(texto || '').split(/[,;]/).filter(function (x) { return x.trim(); });
    var lista = [];
    partes.forEach(function (parte) {
      var nombre = parte.trim(), visible = '';
      var m = /^\s*([^=]+?)\s*=\s*(.+?)\s*$/.exec(nombre);
      if (m) { nombre = m[1]; visible = m[2]; }
      nombre = nombre.replace(/\.exe$/i, '');
      if (!nombre) return;
      if (!visible) visible = (partes.length === 1 && descripcion) ? descripcion : nombre;
      lista.push({ Nombre: nombre, Descripcion: visible });
    });
    return lista;
  }
  // Texto con las etiquetas de PSADT ([bold] [italic] [accent] [url]) como nodos seguros (sin HTML del usuario).
  function textoFormateado(destino, texto) {
    destino.textContent = '';
    if (!texto) return;
    var rx = /\[url=([^\]]+)\](.+?)\[\/url\]|\[url\](.+?)\[\/url\]|\[(\/?(?:bold|italic|accent))\]/g;
    var neg = 0, cur = 0, acc = 0, pos = 0, m;
    function agregar(t, enlace) {
      if (!t) return;
      var s = document.createElement('span');
      s.textContent = t;
      if (neg > 0) s.style.fontWeight = '700';
      if (cur > 0) s.style.fontStyle = 'italic';
      if (acc > 0) s.className = 'acento';
      if (enlace) s.className = 'url';
      destino.appendChild(s);
    }
    while ((m = rx.exec(texto)) !== null) {
      if (m.index > pos) agregar(texto.substring(pos, m.index), false);
      if (m[4]) {
        switch (m[4]) { case 'bold': neg++; break; case '/bold': neg--; break; case 'italic': cur++; break; case '/italic': cur--; break; case 'accent': acc++; break; case '/accent': acc--; break; }
      } else agregar(m[2] || m[3], true);
      pos = m.index + m[0].length;
    }
    if (pos < texto.length) agregar(texto.substring(pos), false);
  }
  function txtAplazar(v) { v = +v; return v === 0 ? 'No permitir' : v === 1 ? '1 vez' : v + ' veces'; }
  function txtDias(v) { v = +v; return v === 0 ? 'Sin límite' : v === 1 ? '1 día' : v + ' días'; }
  function txtCuenta(v) { v = +v; return v === 0 ? 'Sin cuenta' : v + ' min'; }
  function fechaLimite(dias) {
    var d = new Date(Date.now() + dias * 86400000);
    var p = function (n) { return (n < 10 ? '0' : '') + n; };
    var h = d.getHours() % 12; if (h === 0) h = 12;
    return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + p(h) + ':' + p(d.getMinutes()) + ' ' + (d.getHours() < 12 ? 'a. m.' : 'p. m.');
  }
  function val(id) { return $(id).value.trim(); }

  // ---------------- Datos del formulario ----------------
  function obtenerDatos() {
    return {
      Instalador: estado.instalador, Fabricante: val('txtFabricante'), NombreApp: val('txtNombre'), Version: val('txtVersion'),
      Descripcion: val('txtDescripcion'), Arquitectura: $('cmbArq').value, ArgumentosInstalacion: val('txtArgs'),
      DesinstaladorExe: val('txtDes'), ArgumentosDesinstalacion: val('txtArgsDes'), NombreDeteccion: val('txtDeteccion'),
      ProcesosACerrar: val('txtProcesos'), SiemprePreguntar: $('chkSiempre').checked, VecesAplazar: +$('sldAplazar').value,
      DiasLimite: +$('sldDias').value, CuentaRegresivaMin: +$('sldCuenta').value, MensajeBienvenida: val('txtBienvenida'),
      TextoBotonInstalar: val('txtBotonInstalar'), TextoBotonAplazar: val('txtBotonAplazar'), MensajeFinal: val('txtMensaje'),
      LogoApp: estado.logoApp ? estado.logoApp.ruta : '', LogoAppEnVentanas: $('chkLogoVentanas').checked,
      CopiarCarpetaCompleta: $('chkCarpeta').checked, GenerarIntunewin: $('chkIntunewin').checked,
      CarpetaDestino: estado.destino, IncluirFuente: $('chkCopiarFuente').checked
    };
  }
  // Misma validación que Get-ProblemasDatos en PowerShell.
  function problemas(d, tipo) {
    var e = [], a = [];
    if (!d.Instalador) e.push({ c: 'Instalador', t: 'Elige el instalador.' });
    if (!d.NombreApp) e.push({ c: 'NombreApp', t: 'Escribe el nombre de la aplicación.' });
    if (!d.Version) a.push({ c: 'Version', t: tipo === 'PS1' ? 'Sin versión: se usará 1.0.0 para la marca de detección.' : 'Sin versión: la detección buscará la app solo por su nombre.' });
    if (tipo === 'EXE') {
      if (!d.DesinstaladorExe) e.push({ c: 'DesinstaladorExe', t: 'El EXE necesita la ruta del desinstalador en el equipo.' });
      if (!d.ArgumentosInstalacion) a.push({ c: 'ArgumentosInstalacion', t: 'Sin argumentos silenciosos: si el instalador muestra ventanas, Intune quedará esperando.' });
    }
    if (tipo === 'PS1' && d.DesinstaladorExe && d.Instalador) {
      var base = d.Instalador.substring(0, d.Instalador.lastIndexOf('\\') + 1).toLowerCase();
      if (d.DesinstaladorExe.toLowerCase().indexOf(base) !== 0) e.push({ c: 'DesinstaladorExe', t: 'El script de desinstalación debe estar en la carpeta del script de instalación.' });
    }
    if (!d.Descripcion) a.push({ c: 'Descripcion', t: 'Sin descripción: se escribirá una automática.' });
    return { errores: e, avisos: a };
  }
  function marcarPaso(id, ok, texto) { var el = $(id); el.textContent = (ok ? '✔ ' : '○ ') + texto; el.classList.toggle('ok', !!ok); }
  function validar() {
    var d = obtenerDatos(), r = problemas(d, estado.tipo), hay = !!estado.instalador;
    var mapa = { NombreApp: ['txtNombre', 'msgNombre'], Version: ['txtVersion', 'msgVersion'], DesinstaladorExe: ['txtDes', 'msgDes'], ArgumentosInstalacion: ['txtArgs', 'msgArgs'] };
    Object.keys(mapa).forEach(function (k) { $(mapa[k][0]).classList.remove('error', 'aviso'); $(mapa[k][1]).className = 'msg'; });
    var errores = r.errores.filter(function (x) { return x.c !== 'Instalador'; });
    if (hay) {
      r.avisos.map(function (x) { return [x, 'aviso']; }).concat(errores.map(function (x) { return [x, 'error']; })).forEach(function (p) {
        var m = mapa[p[0].c]; if (!m) return;
        $(m[0]).classList.remove('error', 'aviso'); $(m[0]).classList.add(p[1]);
        $(m[1]).className = 'msg ' + p[1]; $(m[1]).textContent = p[0].t;
      });
    }
    var tiene = function (campos) { return errores.some(function (x) { return campos.indexOf(x.c) >= 0; }); };
    marcarPaso('paso1', hay, '1. Instalador');
    marcarPaso('paso2', hay && !tiene(['NombreApp', 'Version']), '2. Información de la app');
    marcarPaso('paso3', hay && !tiene(['DesinstaladorExe']), '3. Cómo se instala');
    marcarPaso('paso4', hay, '4. Ventanas del usuario');
    var res = $('txtResumen');
    if (!hay) { res.className = 'nota gris'; res.textContent = 'Arrastra un instalador para empezar.'; }
    else if (errores.length) { res.className = 'nota roja'; res.textContent = 'Antes de crear: ' + errores.map(function (x) { return x.t; }).join(' '); }
    else if (r.avisos.length) { res.className = 'nota naranja'; res.textContent = 'Listo para crear. Ten en cuenta: ' + r.avisos.map(function (x) { return x.t; }).join(' '); }
    else { res.className = 'nota verde'; res.textContent = '✔ Todo listo para crear el paquete.'; }
    res.style.margin = '0 0 10px';
    return { errores: errores, avisos: r.avisos, todos: r.errores };
  }

  // ---------------- Vista previa (maqueta) ----------------
  function actualizarVista() {
    if (!estado.listo) return;
    var empresa = val('txtEmpresa') || estado.cfg.Empresa || '';
    var nombre = val('txtNombre') || 'Nombre de la app';
    $('mockTitulo').textContent = (nombre + ' ' + val('txtVersion')).trim();
    $('mockSub').textContent = empresa + ' - Instalación de la aplicación';
    var ini = $('mockIni');
    var src = ($('chkLogoVentanas').checked && estado.logoApp) ? urlImagen('id=' + estado.logoApp.id) : (estado.urlLogoCorp || '');
    if (src !== estado.srcMock || !src) {
      estado.srcMock = src;
      ini.textContent = '';
      if (src) { var img = document.createElement('img'); img.src = src; img.alt = ''; img.onerror = function () { ini.textContent = iniciales(val('txtEmpresa') || estado.cfg.Empresa); }; ini.appendChild(img); }
      else ini.textContent = iniciales(empresa);
    }

    var procesos = listaProcesos(val('txtProcesos'), nombre);
    var veces = +$('sldAplazar').value, dias = +$('sldDias').value, cuenta = +$('sldCuenta').value;
    if (val('txtBienvenida')) textoFormateado($('mockTexto'), val('txtBienvenida'));
    else if (procesos.length) textoFormateado($('mockTexto'), 'Por favor, guarde su trabajo antes de continuar ya que las siguientes aplicaciones se cerrarán automáticamente.');
    else textoFormateado($('mockTexto'), 'Por favor, seleccione Instalar para continuar con la instalación.');

    var lista = $('mockLista'); lista.textContent = '';
    procesos.forEach(function (p) {
      var f = document.createElement('div'); f.className = 'proc';
      var i = document.createElement('i'); i.textContent = String(p.Descripcion).substring(0, 1).toUpperCase();
      var t = document.createElement('div'); var b = document.createElement('b'); b.textContent = p.Descripcion; var s = document.createElement('small'); s.textContent = p.Nombre + '.exe';
      t.appendChild(b); t.appendChild(s); f.appendChild(i); f.appendChild(t); lista.appendChild(f);
    });
    lista.classList.toggle('oculto', procesos.length === 0);

    var filas = $('mockFilas'); filas.textContent = '';
    function fila(icono, texto, valor, resaltar) {
      var f = document.createElement('div'); f.className = 'fila-m';
      var a = document.createElement('span'); a.textContent = icono + '  ' + texto;
      var b = document.createElement('b'); b.textContent = valor; if (resaltar) b.className = 'res';
      f.appendChild(a); f.appendChild(b); filas.appendChild(f);
    }
    if (cuenta > 0 && (procesos.length > 0 || veces > 0)) fila('⏱', 'Cuenta regresiva de inicio automático', '0h ' + cuenta + 'm 0s', false);
    if (veces > 0) fila('💬', 'Aplazamientos restantes', String(veces), false);
    if (veces > 0 && dias > 0) fila('📅', 'Fecha límite de aplazamiento', fechaLimite(dias), true);

    $('mockBtnInstalar').textContent = val('txtBotonInstalar') || (procesos.length ? 'Cerrar Aplicaciones e Instalar' : 'Instalar');
    $('mockBtnAplazar').textContent = val('txtBotonAplazar') || 'Aplazar';
    $('mockBtnAplazar').classList.toggle('oculto', veces === 0);
    $('mockBotones').style.gridTemplateColumns = veces > 0 ? '1fr 1fr' : '1fr';

    var aviso = $('txtAvisoSilencio');
    if (procesos.length > 0 && !$('chkSiempre').checked) aviso.textContent = 'Esta ventana solo aparece si alguna de esas apps está ABIERTA. Si están cerradas, se instala en silencio sin molestar al usuario.';
    else if (procesos.length === 0 && veces === 0 && cuenta === 0) aviso.textContent = 'Sin apps a cerrar ni aplazamiento: el usuario solo verá la barra de progreso.';
    else aviso.textContent = 'Esta ventana aparece cada vez que haya un usuario conectado en el equipo.';
    validar();
  }

  // ---------------- Tipo de instalador ----------------
  function mostrar(id, si) { $(id).classList.toggle('oculto', !si); }
  function aplicarModo(tipo) {
    estado.tipo = tipo;
    $('chipTipo').textContent = tipo; mostrar('chipTipo', !!tipo);
    $('chkCarpeta').disabled = false; $('lblCarpeta').classList.remove('deshabilitado');
    if (tipo === 'MSI') {
      $('txtSubInstalacion').textContent = 'MSI: no hay que escribir nada. Se instala, desinstala y detecta con el propio MSI.';
      mostrar('filaArgs', true); mostrar('filaDes', false); mostrar('filaArgsDes', false); mostrar('filaDet', false);
      $('lblArgs').textContent = 'Propiedades MSI (opcional)';
      $('lblAyudaArgs').textContent = 'Ej: SERVIDOR=srv01 LICENCIA=xyz. El modo silencioso (/qn) ya lo pone PSADT.';
    } else if (tipo === 'EXE') {
      $('txtSubInstalacion').textContent = 'EXE: indica cómo se instala en silencio y dónde queda su desinstalador.';
      mostrar('filaArgs', true); mostrar('filaDes', true); mostrar('filaArgsDes', true); mostrar('filaDet', true);
      $('lblArgs').textContent = 'Argumentos silenciosos (obligatorio)';
      $('lblAyudaArgs').textContent = 'Ej: /S   /silent   /quiet   /VERYSILENT /NORESTART  (prueba setup.exe /? para verlos)';
      $('lblDes').textContent = 'Desinstalador en el equipo (obligatorio)';
      $('lblAyudaDes').textContent = 'Ruta después de instalar. Ej: $env:ProgramFiles\\MiApp\\uninstall.exe';
      mostrar('btnDes', false);
    } else if (tipo === 'PS1') {
      $('txtSubInstalacion').textContent = 'Script PS1: se copia toda su carpeta. Tu script debe terminar con «exit 0» si todo salió bien.';
      mostrar('filaArgs', true); mostrar('filaDes', true); mostrar('filaArgsDes', true); mostrar('filaDet', false);
      $('lblArgs').textContent = 'Argumentos del script (opcional)';
      $('lblAyudaArgs').textContent = 'Se pasan al script tal cual. Ej: -Modo Completo';
      $('lblDes').textContent = 'Script de desinstalación (opcional)';
      $('lblAyudaDes').textContent = 'Debe estar en la misma carpeta. Si lo dejas vacío, al desinstalar solo se quita la marca de registro.';
      mostrar('btnDes', true);
      $('chkCarpeta').checked = true; $('chkCarpeta').disabled = true; $('lblCarpeta').classList.add('deshabilitado');
    } else {
      $('txtSubInstalacion').textContent = 'Primero elige un instalador.';
      mostrar('filaArgs', false); mostrar('filaDes', false); mostrar('filaArgsDes', false); mostrar('filaDet', false);
    }
  }
  function mostrarLogoApp(logo, info) {
    var caja = $('logoAppCaja'); caja.textContent = '';
    if (logo && logo.ruta) {
      estado.logoApp = logo;
      var img = document.createElement('img'); img.src = urlImagen('id=' + logo.id); img.alt = ''; caja.appendChild(img);
      $('txtLogoAppInfo').textContent = info || logo.ruta.split('\\').pop();
      $('chkLogoVentanas').disabled = false; $('lblLogoVentanas').classList.remove('deshabilitado');
    } else {
      estado.logoApp = null;
      var s = document.createElement('span'); s.textContent = iniciales(val('txtNombre') || 'App'); caja.appendChild(s);
      $('txtLogoAppInfo').textContent = 'Sin logo: se generan las iniciales con el color de la empresa.';
      $('chkLogoVentanas').checked = false; $('chkLogoVentanas').disabled = true; $('lblLogoVentanas').classList.add('deshabilitado');
    }
    actualizarVista();
  }
  function cargarInstalador(ruta, nombreMostrado) {
    esperar('Leyendo el instalador…', ruta);
    return api('/api/leer', { ruta: ruta }).then(function (d) {
      finEspera();
      if (!d.ok) { aplicarModo(''); return error('No se pudo leer el instalador', d); }
      var info = d.info;
      estado.instalador = d.ruta;
      $('zonaIco').textContent = '✔';
      $('zonaT1').textContent = nombreMostrado || d.ruta.split('\\').pop();
      $('zonaT2').textContent = d.ruta.substring(0, d.ruta.lastIndexOf('\\')) + '   (clic para cambiar)';
      $('txtRuta').value = d.ruta;
      $('txtNombre').value = info.Nombre || '';
      $('txtVersion').value = info.Version || '';
      $('txtFabricante').value = info.Fabricante || estado.cfg.Empresa || '';
      $('txtDeteccion').value = info.Nombre || '';
      $('txtArgs').value = ''; $('txtDes').value = info.Desinstalador || ''; $('txtArgsDes').value = '';
      $('chkCarpeta').checked = false;
      aplicarModo(info.Tipo);
      if (d.subido) { $('chkCarpeta').checked = false; $('chkCarpeta').disabled = true; $('lblCarpeta').classList.add('deshabilitado'); }
      if (info.Tipo === 'MSI' && info.CodigoProducto) $('txtTipoDetalle').textContent = 'ProductCode ' + info.CodigoProducto;
      else if (info.Tipo === 'PS1' && info.Desinstalador) $('txtTipoDetalle').textContent = 'Script de desinstalación detectado: ' + info.Desinstalador.split('\\').pop();
      else $('txtTipoDetalle').textContent = d.subido ? 'Archivo copiado desde el navegador.' : '';
      if (d.icono) mostrarLogoApp(d.icono, 'Ícono extraído del EXE. Puedes cambiarlo.'); else mostrarLogoApp(null);
      actualizarVista();
    });
  }
  function elegirInstalador() {
    if (estado.ocupado) return;
    esperar('Elige el instalador', 'Se abrió la ventana de Windows para buscar el archivo. Si no la ves, búscala en la barra de tareas.');
    api('/api/buscar', { clase: 'instalador', inicial: estado.instalador }).then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo abrir el buscador', d);
      if (d.ruta) cargarInstalador(d.ruta);
    });
  }

  // ---------------- Eventos: Crear paquete ----------------
  var zona = $('zona');
  zona.onclick = elegirInstalador;
  zona.onkeydown = function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); elegirInstalador(); } };
  ['dragenter', 'dragover'].forEach(function (n) { zona.addEventListener(n, function (e) { e.preventDefault(); zona.classList.add('encima'); }); });
  ['dragleave', 'drop'].forEach(function (n) { zona.addEventListener(n, function (e) { e.preventDefault(); zona.classList.remove('encima'); }); });
  zona.addEventListener('drop', function (e) {
    var f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    if (!/\.(msi|exe|ps1)$/i.test(f.name)) { dialogo('Archivo no válido', 'Solo se admiten instaladores .msi, .exe o scripts .ps1.'); return; }
    esperar('Copiando el instalador…', f.name + ' (' + Math.round(f.size / 1048576) + ' MB)');
    subirArchivo(f, 'instalador').then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo copiar el instalador', d);
      return cargarInstalador(d.ruta, f.name);
    });
  });
  // Evita que soltar un archivo fuera de la zona abra el archivo en el navegador.
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) { e.preventDefault(); });

  $('btnUsarRuta').onclick = function () {
    var r = val('txtRuta').replace(/^"+|"+$/g, '');
    if (!r) { dialogo('Falta la ruta', 'Pega la ruta completa del instalador, por ejemplo D:\\Fuentes\\MiApp\\setup.msi'); return; }
    cargarInstalador(r);
  };
  $('txtRuta').onkeydown = function (e) { if (e.key === 'Enter') $('btnUsarRuta').onclick(); };

  $('btnDes').onclick = function () {
    esperar('Elige el script de desinstalación', 'Se abrió la ventana de Windows para buscar el archivo.');
    api('/api/buscar', { clase: 'ps1', inicial: estado.instalador }).then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo abrir el buscador', d);
      if (d.ruta) { $('txtDes').value = d.ruta; validar(); }
    });
  };
  $('btnLogoApp').onclick = function () {
    esperar('Elige el logo de la app', 'Se abrió la ventana de Windows para buscar la imagen.');
    api('/api/buscar', { clase: 'imagen' }).then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo abrir la imagen', d);
      if (d.ruta) mostrarLogoApp({ ruta: d.ruta, id: d.id }, '');
    });
  };
  $('logoAppCaja').addEventListener('drop', function (e) {
    var f = e.dataTransfer.files && e.dataTransfer.files[0]; if (!f) return;
    subirArchivo(f, 'imagen').then(function (d) { if (!d.ok) return error('Logo de la app', d); mostrarLogoApp({ ruta: d.ruta, id: d.id }, f.name); });
  });
  $('btnQuitarLogoApp').onclick = function () { mostrarLogoApp(null); };

  ['txtNombre', 'txtVersion', 'txtProcesos', 'txtBienvenida', 'txtBotonInstalar', 'txtBotonAplazar'].forEach(function (id) { $(id).addEventListener('input', actualizarVista); });
  ['txtFabricante', 'txtDescripcion', 'txtArgs', 'txtDes'].forEach(function (id) { $(id).addEventListener('input', validar); });
  $('txtNombre').addEventListener('input', function () { if (!estado.logoApp) $('logoAppCaja').textContent = iniciales(val('txtNombre') || 'App'); });
  ['chkLogoVentanas', 'chkSiempre'].forEach(function (id) { $(id).addEventListener('change', actualizarVista); });
  $('sldAplazar').addEventListener('input', function () { $('lblAplazar').textContent = txtAplazar(this.value); $('sldDias').disabled = (+this.value === 0); actualizarVista(); });
  $('sldDias').addEventListener('input', function () { $('lblDias').textContent = txtDias(this.value); actualizarVista(); });
  $('sldCuenta').addEventListener('input', function () { $('lblCuenta').textContent = txtCuenta(this.value); actualizarVista(); });

  function mostrarDestino() {
    $('txtDestino').textContent = estado.destino ? estado.destino : estado.salidaPorDefecto + '  (por defecto)';
    $('txtDestino').title = estado.destino || estado.salidaPorDefecto;
  }
  $('btnDestino').onclick = function () {
    esperar('Elige la carpeta destino', 'Se abrió la ventana de Windows para elegir la carpeta.');
    api('/api/buscar', { clase: 'carpeta', inicial: estado.destino }).then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo abrir el buscador', d);
      if (!d.ruta) return;
      api('/api/destino', { ruta: d.ruta }).then(function (r) { if (!r.ok) return error('Carpeta destino', r); estado.destino = d.ruta; mostrarDestino(); });
    });
  };
  $('btnDestinoDefecto').onclick = function () {
    api('/api/destino', { ruta: '' }).then(function (r) { if (!r.ok) return error('Carpeta destino', r); estado.destino = ''; mostrarDestino(); });
  };

  $('btnCrear').onclick = function () {
    if (estado.ocupado) return;
    if (!estado.instalador) { dialogo('Falta el instalador', 'Primero arrastra, busca o pega la ruta del instalador.'); return; }
    var v = validar();
    if (v.errores.length) { dialogo('Faltan datos', 'Corrige esto antes de crear el paquete (está marcado en rojo):\n\n' + v.errores.map(function (x) { return '• ' + x.t; }).join('\n')); return; }
    esperar('Creando el paquete…', 'Puede tardar varios minutos si el instalador es grande o se genera el .intunewin.');
    $('pie').classList.add('abierto');
    api('/api/crear', obtenerDatos()).then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo crear el paquete', d);
      estado.ultimaSalida = d.resultado.RutaSalida;
      $('btnAbrir').disabled = false;
      dialogo('Listo', 'Paquete creado en:\n' + d.resultado.RutaSalida + '\n\nAhí están el .intunewin, FichaIntune.txt, Deteccion.ps1, Descripcion.txt y Logo_Intune.png.',
        [{ t: 'Cerrar', v: false }, { t: 'Abrir la carpeta', v: true, p: true }]).then(function (abrir) { if (abrir) abrirCarpeta(estado.ultimaSalida); });
    });
  };
  $('btnVista').onclick = function () {
    if (estado.ocupado) return;
    esperar('Preparando las ventanas reales…', 'No se instala nada. Las ventanas de PSADT aparecen en unos segundos.');
    api('/api/vista', obtenerDatos()).then(function (d) { finEspera(); if (!d.ok) return error('No se pudo abrir la vista previa', d); });
  };
  function abrirCarpeta(ruta) { api('/api/abrir', { ruta: ruta }).then(function (d) { if (!d.ok) error('Abrir carpeta', d); }); }
  $('btnAbrir').onclick = function () { if (estado.ultimaSalida) abrirCarpeta(estado.ultimaSalida); };

  // ---------------- Paquetes creados ----------------
  var paquetes = [];
  function pintarPaquetes() {
    var filtro = val('txtBuscarPaquete').toLowerCase(), cont = $('listaPaquetes'), n = 0;
    cont.textContent = '';
    paquetes.forEach(function (p) {
      var buscar = [p.NombrePaquete, p.NombreApp, p.Version, p.Fabricante].join(' ').toLowerCase();
      if (filtro && buscar.indexOf(filtro) < 0) return;
      n++;
      var t = document.createElement('div'); t.className = 'tarjeta paquete';
      var tipo = document.createElement('div'); tipo.className = 'tipo'; tipo.textContent = p.Tipo || 'PKG';
      var tx = document.createElement('div'); tx.style.minWidth = '0';
      var n1 = document.createElement('div'); n1.className = 'n'; n1.textContent = p.NombreApp ? (p.NombreApp + ' ' + p.Version).trim() : p.NombrePaquete;
      var det = [p.Fecha]; if (p.Fabricante) det.push(p.Fabricante); if (p.Usuario) det.push('por ' + p.Usuario);
      if (p.TieneIntunewin) det.push('.intunewin listo'); else if (p.ExisteSalida) det.push('sin .intunewin');
      if (!p.ExisteFuente) det.push('carpeta fuente borrada');
      var n2 = document.createElement('div'); n2.className = 'd'; n2.textContent = det.join('  ·  ');
      var n3 = document.createElement('div'); n3.className = 'c'; n3.textContent = p.NombrePaquete;
      tx.appendChild(n1); tx.appendChild(n2); tx.appendChild(n3);
      var acc = document.createElement('div'); acc.className = 'acc';
      [['📂 Salida', p.RutaSalida, p.ExisteSalida], ['🛠 Fuente', p.RutaPaquete, p.ExisteFuente]].forEach(function (b) {
        var bt = document.createElement('button'); bt.className = 'btn'; bt.textContent = b[0]; bt.title = b[1] || ''; bt.disabled = !b[2];
        bt.onclick = function () { abrirCarpeta(b[1]); }; acc.appendChild(bt);
      });
      t.appendChild(tipo); t.appendChild(tx); t.appendChild(acc); cont.appendChild(t);
    });
    if (!paquetes.length) { var v = document.createElement('div'); v.className = 'ayuda'; v.style.margin = '20px 4px'; v.textContent = 'Todavía no hay paquetes. Crea el primero en «Crear paquete».'; cont.appendChild(v); }
    $('txtPaquetesResumen').textContent = 'Mostrando ' + n + ' de ' + paquetes.length + ' paquetes, del más reciente al más antiguo.';
  }
  function cargarPaquetes() { api('/api/paquetes').then(function (d) { if (!d.ok) return; paquetes = d.paquetes || []; pintarPaquetes(); }); }
  $('txtBuscarPaquete').addEventListener('input', pintarPaquetes);
  $('btnRefrescarPaquetes').onclick = cargarPaquetes;

  // ---------------- Plantilla corporativa ----------------
  function marcarPendiente() { if (estado.listo) mostrar('bdPendiente', true); }
  function ponerSrc(img, url) {
    img.style.visibility = 'hidden';
    img.onload = function () { img.style.visibility = 'visible'; };
    img.onerror = function () { img.style.visibility = 'hidden'; };
    img.src = url;
  }
  function imagenEn(caja, src) {
    caja.textContent = '';
    if (!src) return;
    var img = document.createElement('img'); img.src = src; img.alt = ''; img.onerror = function () { caja.textContent = ''; }; caja.appendChild(img);
  }
  function cargarPlantilla(cfg) {
    estado.listo = false;
    estado.cfg = cfg;
    $('txtEmpresa').value = cfg.Empresa || '';
    $('txtColor').value = cfg.ColorAcento || '#2C2A29';
    if (hexValido(cfg.ColorAcento)) $('selColor').value = cfg.ColorAcento;
    aplicarAcento(cfg.ColorAcento || '#2C2A29');
    $('cmbEstilo').value = cfg.EstiloDialogo === 'Classic' ? 'Classic' : 'Fluent';
    $('cmbIdioma').value = (cfg.Idioma === 'es' || cfg.Idioma === 'en') ? cfg.Idioma : 'auto';
    $('txtLogs').value = cfg.RutaLogs || '';
    $('sldAplazarDef').value = Math.min(10, +cfg.VecesAplazar || 0);
    $('sldDiasDef').value = Math.min(30, +cfg.DiasLimite || 0);
    $('sldCuentaDef').value = Math.min(60, +cfg.CuentaRegresivaMin || 0);
    $('lblAplazarDef').textContent = txtAplazar($('sldAplazarDef').value);
    $('lblDiasDef').textContent = txtDias($('sldDiasDef').value);
    $('lblCuentaDef').textContent = txtCuenta($('sldCuentaDef').value);
    $('txtBienvenidaDef').value = cfg.MensajeBienvenida || '';
    $('txtMensajeDef').value = cfg.MensajeFinal || '';
    $('txtBotonInstalarDef').value = cfg.TextoBotonInstalar || '';
    $('txtBotonAplazarDef').value = cfg.TextoBotonAplazar || '';
    estado.logoPendiente = null; estado.bannerPendiente = null;
    imagenEn($('logoEmpCaja'), cfg.Logo ? urlImagen('tipo=logoEmp') : '');
    $('txtLogoEmpInfo').textContent = cfg.Logo ? 'Logo guardado: ' + cfg.Logo : 'Sin logo: se usan las iniciales de la empresa.';
    imagenEn($('bannerEmpCaja'), urlImagen(cfg.Banner ? 'tipo=bannerEmp' : 'tipo=bannerCorp'));
    $('txtBannerEmpInfo').textContent = cfg.Banner ? 'Banner propio: ' + cfg.Banner : 'Generado automáticamente con el logo, el color y el nombre de la empresa.';
    estado.urlLogoCorp = urlImagen('tipo=logoCorp');
    ponerSrc($('imgIconoPrev'), estado.urlLogoCorp);
    ponerSrc($('imgCabecera'), estado.urlLogoCorp);
    ponerSrc($('imgBannerPrev'), urlImagen('tipo=bannerCorp'));
    mostrar('bdPendiente', false);
    // Valores por defecto en «Crear paquete»
    $('sldAplazar').value = Math.min(10, +cfg.VecesAplazar || 0);
    $('sldDias').value = Math.min(30, +cfg.DiasLimite || 0);
    $('sldCuenta').value = Math.min(60, +cfg.CuentaRegresivaMin || 0);
    $('lblAplazar').textContent = txtAplazar($('sldAplazar').value);
    $('lblDias').textContent = txtDias($('sldDias').value);
    $('lblCuenta').textContent = txtCuenta($('sldCuenta').value);
    $('sldDias').disabled = (+$('sldAplazar').value === 0);
    $('txtBienvenida').value = cfg.MensajeBienvenida || '';
    $('txtBotonInstalar').value = cfg.TextoBotonInstalar || '';
    $('txtBotonAplazar').value = cfg.TextoBotonAplazar || '';
    $('txtMensaje').value = cfg.MensajeFinal || '';
    estado.destino = cfg.CarpetaDestino || '';
    mostrarDestino();
    estado.listo = true;
    actualizarVista();
  }
  ['#FDDA24', '#2C2A29', '#0F4C81', '#0078D4', '#00A651', '#E4002B', '#6941C6'].forEach(function (hex) {
    var m = document.createElement('button'); m.className = 'muestra'; m.style.background = hex; m.title = hex;
    m.onclick = function () { $('txtColor').value = hex; $('txtColor').dispatchEvent(new Event('input')); };
    $('paleta').appendChild(m);
  });
  $('txtColor').addEventListener('input', function () {
    var h = val('txtColor');
    if (hexValido(h)) { h = '#' + h.replace('#', ''); $('selColor').value = h; aplicarAcento(h); $('msgColor').className = 'msg'; marcarPendiente(); actualizarVista(); }
    else { $('msgColor').className = 'msg error'; $('msgColor').textContent = 'Usa el formato #RRGGBB (ej. #2C2A29).'; }
  });
  $('selColor').addEventListener('input', function () { $('txtColor').value = this.value.toUpperCase(); $('txtColor').dispatchEvent(new Event('input')); });
  function elegirImagenPlantilla(cual) {
    esperar(cual === 'logo' ? 'Elige el logo de la empresa' : 'Elige el banner', 'Se abrió la ventana de Windows para buscar la imagen.');
    api('/api/buscar', { clase: cual === 'logo' ? 'imagen' : 'banner' }).then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo abrir la imagen', d);
      if (d.ruta) ponerImagenPlantilla(cual, d, d.ruta.split('\\').pop());
    });
  }
  function ponerImagenPlantilla(cual, d, nombre) {
    if (cual === 'logo') {
      estado.logoPendiente = d.ruta;
      imagenEn($('logoEmpCaja'), urlImagen('id=' + d.id));
      $('txtLogoEmpInfo').textContent = nombre + ' (sin guardar)';
    } else {
      estado.bannerPendiente = d.ruta;
      imagenEn($('bannerEmpCaja'), urlImagen('id=' + d.id));
      var txt = nombre + (d.ancho ? ' (' + d.ancho + ' x ' + d.alto + ' px)' : '') + ' - sin guardar';
      if (d.ancho && Math.abs(d.ancho / d.alto - 7.2) / 7.2 > 0.15) txt += '. Proporción distinta a 7,2:1: se centrará sobre fondo blanco.';
      $('txtBannerEmpInfo').textContent = txt;
    }
    marcarPendiente();
  }
  $('btnLogo').onclick = function () { elegirImagenPlantilla('logo'); };
  $('btnBanner').onclick = function () { elegirImagenPlantilla('banner'); };
  [['logoEmpCaja', 'logo'], ['bannerEmpCaja', 'banner']].forEach(function (par) {
    $(par[0]).addEventListener('drop', function (e) {
      var f = e.dataTransfer.files && e.dataTransfer.files[0]; if (!f) return;
      subirArchivo(f, 'imagen').then(function (d) { if (!d.ok) return error('Imagen', d); ponerImagenPlantilla(par[1], d, f.name); });
    });
  });
  $('btnQuitarLogo').onclick = function () { estado.logoPendiente = ''; imagenEn($('logoEmpCaja'), ''); $('txtLogoEmpInfo').textContent = 'Se usarán las iniciales (sin guardar)'; marcarPendiente(); };
  $('btnQuitarBanner').onclick = function () { estado.bannerPendiente = ''; imagenEn($('bannerEmpCaja'), ''); $('txtBannerEmpInfo').textContent = 'Se generará automáticamente (sin guardar)'; marcarPendiente(); };
  ['txtEmpresa', 'txtLogs', 'txtBienvenidaDef', 'txtMensajeDef', 'txtBotonInstalarDef', 'txtBotonAplazarDef'].forEach(function (id) { $(id).addEventListener('input', marcarPendiente); });
  $('txtEmpresa').addEventListener('input', actualizarVista);
  ['cmbEstilo', 'cmbIdioma'].forEach(function (id) { $(id).addEventListener('change', marcarPendiente); });
  $('sldAplazarDef').addEventListener('input', function () { $('lblAplazarDef').textContent = txtAplazar(this.value); marcarPendiente(); });
  $('sldDiasDef').addEventListener('input', function () { $('lblDiasDef').textContent = txtDias(this.value); marcarPendiente(); });
  $('sldCuentaDef').addEventListener('input', function () { $('lblCuentaDef').textContent = txtCuenta(this.value); marcarPendiente(); });
  $('btnGuardar').onclick = function () {
    if (estado.ocupado) return;
    var valores = {
      Empresa: val('txtEmpresa'), ColorAcento: val('txtColor'), EstiloDialogo: $('cmbEstilo').value, Idioma: $('cmbIdioma').value,
      RutaLogs: val('txtLogs'), VecesAplazar: +$('sldAplazarDef').value, DiasLimite: +$('sldDiasDef').value, CuentaRegresivaMin: +$('sldCuentaDef').value,
      MensajeBienvenida: val('txtBienvenidaDef'), MensajeFinal: val('txtMensajeDef'), TextoBotonInstalar: val('txtBotonInstalarDef'), TextoBotonAplazar: val('txtBotonAplazarDef')
    };
    esperar('Guardando la plantilla…', 'Se regeneran el logo y el banner.');
    api('/api/plantilla', { valores: valores, logo: estado.logoPendiente, banner: estado.bannerPendiente }).then(function (d) {
      finEspera();
      if (!d.ok) return error('No se pudo guardar la plantilla', d);
      cargarPlantilla(d.cfg);
      dialogo('Plantilla corporativa', 'Plantilla guardada.\n\nSe aplicará a los paquetes que crees desde ahora. Los paquetes que ya existían no cambian: vuelve a crearlos.');
    });
  };
  $('btnAbrirRecursos').onclick = function () { abrirCarpeta('Recursos'); };

  // ---------------- Cabecera ----------------
  $('btnTema').onclick = function () {
    var r = document.documentElement, oscuro = r.dataset.theme ? r.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
    r.dataset.theme = oscuro ? 'light' : 'dark';
    try { localStorage.setItem('fabricaTema', r.dataset.theme); } catch (e) { }
  };
  try { var tema = localStorage.getItem('fabricaTema'); if (tema) document.documentElement.dataset.theme = tema; } catch (e) { }
  $('btnSalir').onclick = function () {
    dialogo('Cerrar la Fábrica', '¿Detener la Fábrica web? Se cierra la consola y esta página deja de funcionar.', [{ t: 'Cancelar', v: false }, { t: 'Cerrar Fábrica', v: true, p: true }]).then(function (si) {
      if (!si) return;
      api('/api/salir', {}).then(function () {
        document.body.innerHTML = '<div style="margin:80px auto;max-width:480px;text-align:center;font-family:Segoe UI,sans-serif"><h2>La Fábrica se cerró</h2><p>Ya puedes cerrar esta pestaña. Para volver a usarla abre <b>Crear-Paquete-Web.cmd</b>.</p></div>';
      });
    });
  };

  // ---------------- Atajos ----------------
  document.addEventListener('keydown', function (e) {
    if (!e.ctrlKey) return;
    if (e.key === 'o' || e.key === 'O') { e.preventDefault(); irA('crear'); elegirInstalador(); }
    else if (e.key === 'Enter' && estado.pagina === 'crear') { e.preventDefault(); $('btnCrear').onclick(); }
  });

  // ---------------- Inicio ----------------
  estado.pagina = 'crear';
  aplicarModo('');
  mostrarLogoApp(null);
  api('/api/estado').then(function (d) {
    if (!d.ok) { $('txtVersionPsadt').textContent = 'Sin conexión'; return error('No se pudo conectar con la Fábrica', d); }
    $('txtVersionPsadt').textContent = d.versionPsadt ? 'PSADT ' + d.versionPsadt : 'PSADT no cargado';
    estado.salidaPorDefecto = d.salidaPorDefecto;
    cargarPlantilla(d.cfg);
    if (d.errorInicio) {
      $('pie').classList.add('abierto');
      ['btnCrear', 'btnVista', 'btnGuardar'].forEach(function (id) { $(id).disabled = true; });
      error('No se pudo preparar la Fábrica', { error: d.errorInicio });
    } else if (!d.tieneIntuneWinAppUtil) {
      $('chkIntunewin').checked = false;
    }
  });
})();
</script>
</body>
</html>
'@

function New-TokenWeb {
    $bytes = New-Object byte[] 24
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    return (($bytes | ForEach-Object { $_.ToString('x2') }) -join '')
}

function Send-RespuestaWeb {
    param(
        [Parameter(Mandatory)]$Flujo,
        [int]$Codigo = 200,
        [string]$TipoContenido = 'application/json; charset=utf-8',
        [byte[]]$Cuerpo = $null,
        [string[]]$Cabeceras = @()
    )
    $textos = @{ 200 = 'OK'; 204 = 'No Content'; 400 = 'Bad Request'; 403 = 'Forbidden'; 404 = 'Not Found'; 405 = 'Method Not Allowed'; 500 = 'Internal Server Error' }
    $estadoTexto = 'OK'
    if ($textos.ContainsKey($Codigo)) { $estadoTexto = $textos[$Codigo] }
    $largo = 0
    if ($null -ne $Cuerpo) { $largo = $Cuerpo.Length }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("HTTP/1.1 $Codigo $estadoTexto`r`n")
    [void]$sb.Append("Content-Type: $TipoContenido`r`n")
    [void]$sb.Append("Content-Length: $largo`r`n")
    [void]$sb.Append("Connection: close`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nReferrer-Policy: no-referrer`r`nX-Frame-Options: DENY`r`n")
    foreach ($linea in $Cabeceras) { [void]$sb.Append($linea + "`r`n") }
    [void]$sb.Append("`r`n")
    $cab = [System.Text.Encoding]::ASCII.GetBytes($sb.ToString())
    $Flujo.Write($cab, 0, $cab.Length)
    if ($largo -gt 0) { $Flujo.Write($Cuerpo, 0, $largo) }
    $Flujo.Flush()
}

function Send-JsonWeb {
    param([Parameter(Mandatory)]$Flujo, [Parameter(Mandatory)][System.Collections.IDictionary]$Objeto, [int]$Codigo = 200)
    $json = ConvertTo-Json -InputObject $Objeto -Depth 8 -Compress
    Send-RespuestaWeb -Flujo $Flujo -Codigo $Codigo -Cuerpo $script:Utf8SinBom.GetBytes($json)
}

function Read-PeticionWeb {
    # Lee la línea de petición y las cabeceras. Devuelve @{ Metodo; Ruta; Query; Cabeceras; Sobrante },
    # donde Sobrante son los bytes del cuerpo que ya llegaron junto con las cabeceras.
    param([Parameter(Mandatory)]$Flujo)
    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $acumulado = New-Object System.IO.MemoryStream
    $buffer = New-Object byte[] 16384
    $fin = -1
    $texto = ''
    while ($fin -lt 0) {
        $n = $Flujo.Read($buffer, 0, $buffer.Length)
        if ($n -le 0) { return $null }
        $acumulado.Write($buffer, 0, $n)
        $texto = $latin1.GetString($acumulado.GetBuffer(), 0, [int]$acumulado.Length)
        $fin = $texto.IndexOf("`r`n`r`n", [System.StringComparison]::Ordinal)
        if ($fin -lt 0 -and $acumulado.Length -gt 65536) { throw 'Cabeceras HTTP demasiado grandes.' }
    }
    $lineas = $texto.Substring(0, $fin) -split "`r`n"
    $partes = $lineas[0] -split ' '
    if ($partes.Count -lt 3) { throw 'Petición HTTP no válida.' }
    $cabeceras = @{}
    for ($i = 1; $i -lt $lineas.Count; $i++) {
        $p = $lineas[$i].IndexOf(':')
        if ($p -gt 0) { $cabeceras[$lineas[$i].Substring(0, $p).Trim().ToLowerInvariant()] = $lineas[$i].Substring($p + 1).Trim() }
    }
    $destino = $partes[1]
    $ruta = $destino
    $query = @{}
    $q = $destino.IndexOf('?')
    if ($q -ge 0) {
        $ruta = $destino.Substring(0, $q)
        foreach ($par in $destino.Substring($q + 1).Split('&')) {
            if (-not $par) { continue }
            $kv = $par.Split([char[]]@('='), 2)
            $clave = [System.Uri]::UnescapeDataString($kv[0].Replace('+', ' '))
            $valor = ''
            if ($kv.Count -gt 1) { $valor = [System.Uri]::UnescapeDataString($kv[1].Replace('+', ' ')) }
            $query[$clave] = $valor
        }
    }
    $todo = $acumulado.ToArray()
    $inicioCuerpo = $fin + 4
    $sobrante = New-Object byte[] ($todo.Length - $inicioCuerpo)
    if ($sobrante.Length -gt 0) { [Array]::Copy($todo, $inicioCuerpo, $sobrante, 0, $sobrante.Length) }
    return @{ Metodo = $partes[0].ToUpperInvariant(); Ruta = $ruta; Query = $query; Cabeceras = $cabeceras; Sobrante = $sobrante }
}

function Read-CuerpoWeb {
    # Lee el cuerpo (Content-Length). Con -RutaArchivo lo escribe directo a disco (instaladores grandes).
    param([Parameter(Mandatory)]$Flujo, [Parameter(Mandatory)][hashtable]$Peticion, [long]$Maximo = 10MB, [string]$RutaArchivo)
    $largo = 0L
    if ($Peticion.Cabeceras.ContainsKey('content-length')) { $largo = [long]$Peticion.Cabeceras['content-length'] }
    if ($largo -gt $Maximo) { throw ('Los datos son demasiado grandes (máximo {0:N0} MB).' -f ($Maximo / 1MB)) }
    if ($RutaArchivo) { $destino = [System.IO.File]::Create($RutaArchivo) } else { $destino = New-Object System.IO.MemoryStream }
    try {
        $sobrante = $Peticion.Sobrante
        $usar = [Math]::Min([long]$sobrante.Length, $largo)
        if ($usar -gt 0) { $destino.Write($sobrante, 0, [int]$usar) }
        $faltan = $largo - $usar
        $buffer = New-Object byte[] 1048576
        while ($faltan -gt 0) {
            $n = $Flujo.Read($buffer, 0, [int][Math]::Min([long]$buffer.Length, $faltan))
            if ($n -le 0) { throw 'La conexión se cortó mientras se recibían los datos.' }
            $destino.Write($buffer, 0, $n)
            $faltan -= $n
        }
        if ($RutaArchivo) { return }
        return , $destino.ToArray()
    }
    finally { $destino.Dispose() }
}

function Register-ArchivoWeb {
    # Registra un archivo que la página puede ver como imagen (/api/imagen?id=...). Solo se sirven archivos registrados.
    param([Parameter(Mandatory)][hashtable]$Web, [Parameter(Mandatory)][string]$Ruta)
    $id = [guid]::NewGuid().ToString('N')
    $Web.Archivos[$id] = $Ruta
    return $id
}

function Get-DatosImagenWeb {
    # Ruta, id y tamaño de una imagen para la página. Falla si no es una imagen válida.
    param([Parameter(Mandatory)][hashtable]$Web, [Parameter(Mandatory)][string]$Ruta)
    $ancho = 0; $alto = 0
    $img = $null
    try { $img = Get-ImagenSinBloqueo $Ruta } catch { $img = $null }
    if ($null -ne $img) { $ancho = $img.Width; $alto = $img.Height; $img.Dispose() }
    elseif ([System.IO.Path]::GetExtension($Ruta).ToLowerInvariant() -ne '.ico') { throw 'No se pudo abrir la imagen.' }
    return @{ ok = $true; ruta = $Ruta; id = (Register-ArchivoWeb -Web $Web -Ruta $Ruta); ancho = $ancho; alto = $alto }
}

function Show-DialogoWeb {
    # Abre el buscador de archivos/carpetas de Windows encima del navegador (la página no puede ver rutas locales).
    param([string]$Filtro, [string]$Titulo, [string]$Inicial, [switch]$Carpeta)
    $dlg = $null
    $duenio = New-Object System.Windows.Forms.Form
    $duenio.TopMost = $true
    $duenio.ShowInTaskbar = $false
    $duenio.FormBorderStyle = 'None'
    $duenio.Opacity = 0
    $duenio.StartPosition = 'CenterScreen'
    $duenio.Size = New-Object System.Drawing.Size(1, 1)
    try {
        $duenio.Show()
        $duenio.Activate()
        $carpetaInicial = ''
        if ($Inicial) {
            if (Test-Path -LiteralPath $Inicial -PathType Container) { $carpetaInicial = $Inicial }
            elseif (Test-Path -LiteralPath $Inicial -PathType Leaf) { $carpetaInicial = Split-Path -Path $Inicial -Parent }
        }
        if ($Carpeta) {
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = $Titulo
            $dlg.ShowNewFolderButton = $true
            if ($carpetaInicial) { $dlg.SelectedPath = $carpetaInicial }
            if ($dlg.ShowDialog($duenio) -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
        }
        else {
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = $Filtro
            $dlg.Title = $Titulo
            if ($carpetaInicial) { $dlg.InitialDirectory = $carpetaInicial }
            if ($dlg.ShowDialog($duenio) -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
        }
        return ''
    }
    finally {
        if ($null -ne $dlg) { $dlg.Dispose() }
        $duenio.Close()
        $duenio.Dispose()
    }
}

function Test-RutaPermitidaWeb {
    # Solo se abren en el Explorador las carpetas de la Fábrica y de los paquetes del historial.
    param([Parameter(Mandatory)][hashtable]$Web, [string]$Ruta)
    if ([string]::IsNullOrWhiteSpace($Ruta)) { return $false }
    try { if (-not [System.IO.Path]::IsPathRooted($Ruta)) { return $false } } catch { return $false }
    $normalizar = { param([string]$r) try { [System.IO.Path]::GetFullPath($r).TrimEnd('\').ToLowerInvariant() } catch { '' } }
    $objetivo = & $normalizar $Ruta
    $permitidas = @($script:Recursos, $script:CarpetaSalida, $script:CarpetaPaquetes, [string]$Web.UltimaSalida, [string](Get-ConfigCorporativa).CarpetaDestino)
    foreach ($p in (Get-HistorialPaquetes)) { $permitidas += [string]$p.RutaSalida; $permitidas += [string]$p.RutaPaquete }
    foreach ($p in $permitidas) { if ($p -and (& $normalizar $p) -eq $objetivo) { return $true } }
    return $false
}

function Open-NavegadorWeb {
    # Abre Edge en modo aplicación (ventana sin barra de direcciones). Si no está, el navegador por defecto.
    param([Parameter(Mandatory)][string]$Url)
    $candidatos = @()
    foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA)) {
        if ($base) { $candidatos += (Join-Path $base 'Microsoft\Edge\Application\msedge.exe') }
    }
    $edge = $candidatos | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    try {
        if ($edge) { Start-Process -FilePath $edge -ArgumentList ('--app="{0}"' -f $Url) }
        else { Start-Process -FilePath $Url }
    }
    catch { Write-Registro -Mensaje "No se pudo abrir el navegador. Abre esta dirección a mano: $Url" -Tipo Aviso }
}

function Invoke-PeticionWeb {
    param([Parameter(Mandatory)][System.Net.Sockets.TcpClient]$Cliente, [Parameter(Mandatory)][hashtable]$Web)
    $flujo = $Cliente.GetStream()
    $flujo.ReadTimeout = 30000
    $flujo.WriteTimeout = 30000
    $pet = Read-PeticionWeb -Flujo $flujo
    if ($null -eq $pet) { return }

    $textoPlano = 'text/plain; charset=utf-8'
    $puerto = $Web.Puerto
    $hostPeticion = ([string]$pet.Cabeceras['host']).ToLowerInvariant()
    if (@("127.0.0.1:$puerto", "localhost:$puerto") -notcontains $hostPeticion) {
        Send-RespuestaWeb -Flujo $flujo -Codigo 403 -TipoContenido $textoPlano -Cuerpo $script:Utf8SinBom.GetBytes('Acceso denegado.')
        return
    }
    if ($pet.Cabeceras.ContainsKey('origin')) {
        $origen = ([string]$pet.Cabeceras['origin']).ToLowerInvariant()
        if (@("http://127.0.0.1:$puerto", "http://localhost:$puerto") -notcontains $origen) {
            Send-RespuestaWeb -Flujo $flujo -Codigo 403 -TipoContenido $textoPlano -Cuerpo $script:Utf8SinBom.GetBytes('Acceso denegado.')
            return
        }
    }
    $tokenPeticion = [string]$pet.Cabeceras['x-fabrica-token']
    if (-not $tokenPeticion) { $tokenPeticion = [string]$pet.Query['t'] }
    $tokenOk = ($tokenPeticion -ceq $Web.Token)

    if ($pet.Ruta -eq '/' -and $pet.Metodo -eq 'GET') {
        if (-not $tokenOk) {
            $aviso = '<!doctype html><meta charset="utf-8"><title>Fábrica de Paquetes</title><body style="font-family:Segoe UI,sans-serif;margin:60px"><h2>Enlace no válido</h2><p>Abre la Fábrica con <b>Crear-Paquete-Web.cmd</b>: el enlace cambia cada vez que se abre.</p></body>'
            Send-RespuestaWeb -Flujo $flujo -Codigo 403 -TipoContenido 'text/html; charset=utf-8' -Cuerpo $script:Utf8SinBom.GetBytes($aviso)
            return
        }
        $csp = "Content-Security-Policy: default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'"
        Send-RespuestaWeb -Flujo $flujo -TipoContenido 'text/html; charset=utf-8' -Cuerpo $script:Utf8SinBom.GetBytes($script:HtmlWeb) -Cabeceras @($csp)
        return
    }
    if ($pet.Ruta -eq '/favicon.ico') { Send-RespuestaWeb -Flujo $flujo -Codigo 204 -TipoContenido $textoPlano; return }
    if (-not $pet.Ruta.StartsWith('/api/')) { Send-RespuestaWeb -Flujo $flujo -Codigo 404 -TipoContenido $textoPlano -Cuerpo $script:Utf8SinBom.GetBytes('No existe.'); return }
    if (-not $tokenOk) { Send-JsonWeb -Flujo $flujo -Codigo 403 -Objeto @{ ok = $false; error = 'Código de sesión no válido. Vuelve a abrir la Fábrica con Crear-Paquete-Web.cmd.' }; return }

    $seqInicio = $script:SecuenciaRegistro
    if ($pet.Ruta -eq '/api/estado') { $seqInicio = 0 }   # al abrir la página se envía el registro completo
    $responder = {
        param([System.Collections.IDictionary]$Objeto, [int]$Codigo = 200)
        $nuevas = New-Object System.Collections.Generic.List[object]
        foreach ($l in $script:RegistroWeb) { if ($l.Seq -gt $seqInicio) { $nuevas.Add(@{ Linea = $l.Linea; Tipo = $l.Tipo; Mensaje = $l.Mensaje }) } }
        $Objeto['registro'] = [object[]]$nuevas.ToArray()
        Send-JsonWeb -Flujo $flujo -Objeto $Objeto -Codigo $Codigo
    }
    $leerJson = {
        $bytes = Read-CuerpoWeb -Flujo $flujo -Peticion $pet -Maximo 10MB
        if ($null -eq $bytes -or $bytes.Length -eq 0) { return $null }
        return (ConvertFrom-Json ($script:Utf8SinBom.GetString($bytes)))
    }

    try {
        switch ("$($pet.Metodo) $($pet.Ruta)") {
            'GET /api/estado' {
                $versionPsadt = ''
                if ($script:VersionPSADT -gt [version]'0.0') { $versionPsadt = [string]$script:VersionPSADT }
                & $responder @{
                    ok                    = $true
                    versionPsadt          = $versionPsadt
                    versionFabrica        = $script:VersionFabrica
                    cfg                   = (Get-ConfigCorporativa)
                    salidaPorDefecto      = $script:CarpetaSalida
                    tieneIntuneWinAppUtil = [bool](Test-Path -LiteralPath (Join-Path $script:Herramientas 'IntuneWinAppUtil.exe'))
                    errorInicio           = [string]$Web.ErrorInicio
                }
            }
            'POST /api/buscar' {
                $cuerpo = & $leerJson
                $clase = [string](Get-PropiedadJson $cuerpo 'clase' '')
                $inicial = [string](Get-PropiedadJson $cuerpo 'inicial' '')
                $filtroImagen = 'Imágenes (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico)|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico'
                $ruta = ''
                if ($clase -eq 'instalador') { $ruta = Show-DialogoWeb -Filtro 'Instaladores (*.msi;*.exe;*.ps1)|*.msi;*.exe;*.ps1' -Titulo 'Selecciona el instalador' -Inicial $inicial }
                elseif ($clase -eq 'ps1') { $ruta = Show-DialogoWeb -Filtro 'Script PowerShell (*.ps1)|*.ps1' -Titulo 'Selecciona el script de desinstalación' -Inicial $inicial }
                elseif ($clase -eq 'imagen') { $ruta = Show-DialogoWeb -Filtro $filtroImagen -Titulo 'Selecciona la imagen' -Inicial $inicial }
                elseif ($clase -eq 'banner') { $ruta = Show-DialogoWeb -Filtro 'Imágenes (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp' -Titulo 'Selecciona el banner (recomendado PNG de 900 x 125 px)' -Inicial $inicial }
                elseif ($clase -eq 'carpeta') { $ruta = Show-DialogoWeb -Carpeta -Titulo 'Elige la carpeta donde quedará el paquete listo (.intunewin, ficha, detección y logo)' -Inicial $inicial }
                else { throw "Tipo de búsqueda no válido: $clase" }
                if ($ruta -and ($clase -eq 'imagen' -or $clase -eq 'banner')) { & $responder (Get-DatosImagenWeb -Web $Web -Ruta $ruta) }
                else { & $responder @{ ok = $true; ruta = [string]$ruta } }
            }
            'POST /api/subir' {
                $clase = [string]$pet.Query['clase']
                $nombre = [System.IO.Path]::GetFileName([string]$pet.Query['nombre'])
                foreach ($caracter in [System.IO.Path]::GetInvalidFileNameChars()) { $nombre = $nombre.Replace([string]$caracter, '_') }
                $extension = [System.IO.Path]::GetExtension($nombre).ToLowerInvariant()
                $permitidas = @()
                if ($clase -eq 'instalador') { $permitidas = @('.msi', '.exe', '.ps1') }
                elseif ($clase -eq 'imagen') { $permitidas = @('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.ico') }
                if ($permitidas -notcontains $extension) { throw "Tipo de archivo no permitido: '$nombre'." }
                $carpeta = Join-Path $Web.Subidas ([guid]::NewGuid().ToString('N'))
                New-Item -Path $carpeta -ItemType Directory -Force | Out-Null
                $destino = Join-Path $carpeta $nombre
                Read-CuerpoWeb -Flujo $flujo -Peticion $pet -Maximo 8GB -RutaArchivo $destino
                Write-Registro -Mensaje ('Archivo recibido desde el navegador: {0} ({1:N1} MB)' -f $nombre, ((Get-Item -LiteralPath $destino).Length / 1MB))
                if ($clase -eq 'imagen') { & $responder (Get-DatosImagenWeb -Web $Web -Ruta $destino) }
                else { & $responder @{ ok = $true; ruta = $destino } }
            }
            'POST /api/leer' {
                $cuerpo = & $leerJson
                $ruta = ([string](Get-PropiedadJson $cuerpo 'ruta' '')).Trim().Trim('"')
                if (-not $ruta) { throw 'Escribe la ruta del instalador.' }
                if (-not [System.IO.Path]::IsPathRooted($ruta)) { throw 'Escribe la ruta completa, por ejemplo D:\Fuentes\MiApp\setup.msi' }
                $ruta = [System.IO.Path]::GetFullPath($ruta)
                $info = Get-InfoInstalador -Ruta $ruta
                $icono = $null
                if ($info.Tipo -eq 'EXE') {
                    $bmp = Get-IconoDeExe -Ruta $ruta
                    if ($null -ne $bmp) {
                        $rutaIcono = Join-Path $Web.Subidas ('Icono_' + [guid]::NewGuid().ToString('N') + '.png')
                        try { $bmp.Save($rutaIcono, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $bmp.Dispose() }
                        $icono = @{ ruta = $rutaIcono; id = (Register-ArchivoWeb -Web $Web -Ruta $rutaIcono) }
                    }
                }
                $subido = $ruta.StartsWith($Web.Subidas.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)
                Write-Registro -Mensaje "Instalador leído: $($info.Nombre) $($info.Version) [$($info.Tipo)]" -Tipo OK
                if ([string]::IsNullOrWhiteSpace($info.Version)) { Write-Registro -Mensaje 'El instalador no trae versión: escríbela a mano.' -Tipo Aviso }
                & $responder @{ ok = $true; ruta = $ruta; info = $info; icono = $icono; subido = $subido }
            }
            'POST /api/crear' {
                $datos = ConvertTo-DatosPaquete (& $leerJson)
                $tipo = switch ([System.IO.Path]::GetExtension([string]$datos.Instalador).ToLowerInvariant()) { '.msi' { 'MSI' } '.exe' { 'EXE' } '.ps1' { 'PS1' } default { '' } }
                $problemas = Get-ProblemasDatos -Datos $datos -Tipo $tipo
                if (@($problemas.Errores).Count -gt 0) {
                    & $responder @{ ok = $false; error = ('Faltan datos: ' + ((@($problemas.Errores) | ForEach-Object { $_.Texto }) -join ' ')) }
                    return
                }
                $resultado = @(New-PaqueteCorporativo -Datos $datos)[-1]
                $Web.UltimaSalida = [string]$resultado.RutaSalida
                & $responder @{ ok = $true; resultado = $resultado }
            }
            'POST /api/vista' {
                $datos = ConvertTo-DatosPaquete (& $leerJson)
                Start-VistaPrevia -Datos $datos
                & $responder @{ ok = $true }
            }
            'POST /api/destino' {
                $cuerpo = & $leerJson
                $ruta = ([string](Get-PropiedadJson $cuerpo 'ruta' '')).Trim()
                if ($ruta -and -not [System.IO.Path]::IsPathRooted($ruta)) { throw 'La carpeta destino debe ser una ruta completa.' }
                $cfgD = Get-ConfigCorporativa
                $cfgD.CarpetaDestino = $ruta
                Save-ConfigCorporativa -Cfg $cfgD
                if ($ruta) { Write-Registro -Mensaje "Los paquetes quedarán en: $ruta" -Tipo OK }
                else { Write-Registro -Mensaje 'Los paquetes quedarán en la carpeta Salida de la Fábrica.' -Tipo OK }
                & $responder @{ ok = $true }
            }
            'POST /api/abrir' {
                $cuerpo = & $leerJson
                $ruta = [string](Get-PropiedadJson $cuerpo 'ruta' '')
                if ($ruta -eq 'Recursos') { $ruta = $script:Recursos }
                if (-not (Test-RutaPermitidaWeb -Web $Web -Ruta $ruta)) { throw 'Por seguridad, desde la página solo se abren las carpetas de la Fábrica y de sus paquetes.' }
                if (-not (Test-Path -LiteralPath $ruta -PathType Container)) { throw "La carpeta ya no existe: $ruta" }
                Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $ruta)
                & $responder @{ ok = $true }
            }
            'GET /api/paquetes' {
                $lista = Get-HistorialPaquetes
                foreach ($p in $lista) { $p['TieneIntunewin'] = [bool]($p.Intunewin -and (Test-Path -LiteralPath $p.Intunewin)) }
                & $responder @{ ok = $true; paquetes = [object[]]$lista }
            }
            'POST /api/plantilla' {
                $cuerpo = & $leerJson
                $valoresObj = Get-PropiedadJson $cuerpo 'valores' $null
                if ($null -eq $valoresObj) { throw 'Faltan los valores de la plantilla.' }
                $valores = @{}
                foreach ($k in @('Empresa', 'ColorAcento', 'EstiloDialogo', 'Idioma', 'RutaLogs', 'VecesAplazar', 'DiasLimite', 'CuentaRegresivaMin', 'MensajeBienvenida', 'MensajeFinal', 'TextoBotonInstalar', 'TextoBotonAplazar')) {
                    $valores[$k] = Get-PropiedadJson $valoresObj $k ''
                }
                $logo = Get-PropiedadJson $cuerpo 'logo' $null
                $banner = Get-PropiedadJson $cuerpo 'banner' $null
                $cfgNueva = @(Save-PlantillaCorporativa -Valores $valores -LogoPendiente $logo -BannerPendiente $banner)[-1]
                & $responder @{ ok = $true; cfg = $cfgNueva }
            }
            'GET /api/imagen' {
                $ruta = ''
                if ($pet.Query.ContainsKey('id')) {
                    if ($Web.Archivos.ContainsKey([string]$pet.Query['id'])) { $ruta = $Web.Archivos[[string]$pet.Query['id']] }
                }
                else {
                    $cfgI = Get-ConfigCorporativa
                    switch ([string]$pet.Query['tipo']) {
                        'logoCorp' { $ruta = Get-RutaLogoCorporativo }
                        'bannerCorp' { $ruta = Get-RutaBannerCorporativo }
                        'logoEmp' { if ([string]$cfgI.Logo) { $ruta = Join-Path $script:Recursos ([string]$cfgI.Logo) } }
                        'bannerEmp' { if ([string]$cfgI.Banner) { $ruta = Join-Path $script:Recursos ([string]$cfgI.Banner) } }
                    }
                }
                $tipos = @{ '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.gif' = 'image/gif'; '.bmp' = 'image/bmp'; '.ico' = 'image/x-icon' }
                $extension = ''
                if ($ruta) { $extension = [System.IO.Path]::GetExtension($ruta).ToLowerInvariant() }
                if (-not $ruta -or -not $tipos.ContainsKey($extension) -or -not (Test-Path -LiteralPath $ruta -PathType Leaf)) {
                    Send-RespuestaWeb -Flujo $flujo -Codigo 404 -TipoContenido $textoPlano
                    return
                }
                Send-RespuestaWeb -Flujo $flujo -TipoContenido $tipos[$extension] -Cuerpo ([System.IO.File]::ReadAllBytes($ruta))
            }
            'POST /api/salir' {
                $Web.Salir = $true
                Write-Registro -Mensaje 'Cerrando la Fábrica web a pedido del usuario.'
                & $responder @{ ok = $true }
            }
            default {
                Send-JsonWeb -Flujo $flujo -Codigo 404 -Objeto @{ ok = $false; error = "No existe: $($pet.Metodo) $($pet.Ruta)" }
            }
        }
    }
    catch {
        $mensaje = $_.Exception.Message
        Write-Registro -Mensaje $mensaje -Tipo Error
        & $responder @{ ok = $false; error = $mensaje } 500
    }
}

function Start-ModoWeb {
    param([int]$Puerto = 0, [switch]$SinNavegador)
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $script:RegistroWeb = New-Object System.Collections.Generic.List[object]
    $web = @{
        Token        = New-TokenWeb
        Puerto       = 0
        Archivos     = @{}
        Subidas      = (Join-Path $env:TEMP ('FabricaWeb_' + [guid]::NewGuid().ToString('N')))
        UltimaSalida = ''
        ErrorInicio  = ''
        Salir        = $false
    }
    New-Item -Path $web.Subidas -ItemType Directory -Force | Out-Null

    $escucha = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Puerto)
    try { $escucha.Start() }
    catch { throw "No se pudo abrir el puerto $Puerto en 127.0.0.1 ($($_.Exception.Message)). Prueba con otro, por ejemplo -Puerto 8765, o no indiques puerto para que se elija uno libre." }
    $web.Puerto = ([System.Net.IPEndPoint]$escucha.LocalEndpoint).Port
    $url = 'http://127.0.0.1:{0}/?t={1}' -f $web.Puerto, $web.Token
    $clientes = New-Object System.Collections.Generic.List[object]

    try {
        try { Initialize-Fabrica }
        catch {
            $web.ErrorInicio = $_.Exception.Message
            Write-Registro -Mensaje $web.ErrorInicio -Tipo Error
        }
        Write-Host ''
        Write-Host '  Fábrica de Paquetes - versión web' -ForegroundColor Cyan
        Write-Host "  Dirección (solo este equipo): $url"
        Write-Host '  NO cierres esta ventana mientras uses la página.' -ForegroundColor Yellow
        Write-Host '  Para terminar: botón «Cerrar Fábrica» en la página, o Ctrl+C aquí.' -ForegroundColor Yellow
        Write-Host ''
        if (-not $SinNavegador) { Open-NavegadorWeb -Url $url }

        # Bucle de un solo hilo. Las conexiones se guardan en una lista y se atienden cuando traen datos:
        # así una conexión abierta «por adelantado» por el navegador no bloquea a las demás.
        while (-not $web.Salir) {
            while ($escucha.Pending()) {
                $tc = $escucha.AcceptTcpClient()
                $tc.NoDelay = $true
                $clientes.Add(@{ Cliente = $tc; Desde = [DateTime]::UtcNow })
            }
            $atendio = $false
            foreach ($item in $clientes.ToArray()) {
                $cerrar = $false
                try {
                    if ($item.Cliente.Available -gt 0) {
                        Invoke-PeticionWeb -Cliente $item.Cliente -Web $web
                        $cerrar = $true
                        $atendio = $true
                    }
                    elseif (-not $item.Cliente.Connected -or ([DateTime]::UtcNow - $item.Desde).TotalSeconds -gt 60) { $cerrar = $true }
                }
                catch {
                    $cerrar = $true
                    Write-Registro -Mensaje "Petición web interrumpida: $($_.Exception.Message)" -Tipo Aviso
                }
                if ($cerrar) {
                    try { $item.Cliente.Close() } catch { }
                    [void]$clientes.Remove($item)
                }
            }
            if (-not $atendio) { Start-Sleep -Milliseconds 30 }
        }
    }
    finally {
        foreach ($item in $clientes) { try { $item.Cliente.Close() } catch { } }
        $escucha.Stop()
        $script:RegistroWeb = $null
        Remove-Item -LiteralPath $web.Subidas -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host 'Fábrica web detenida.'
    }
}
#endregion

#region ===================== Punto de entrada =====================
try {
    if ($SinInterfaz) {
        if ([string]::IsNullOrWhiteSpace($Instalador)) { throw 'Con -SinInterfaz debes indicar -Instalador.' }
        Initialize-Fabrica
        $cfgInicial = Get-ConfigCorporativa
        if (-not $PSBoundParameters.ContainsKey('MensajeFinal')) { $MensajeFinal = [string]$cfgInicial.MensajeFinal }
        if (-not $PSBoundParameters.ContainsKey('MensajeBienvenida')) { $MensajeBienvenida = [string]$cfgInicial.MensajeBienvenida }
        if (-not $PSBoundParameters.ContainsKey('TextoBotonInstalar')) { $TextoBotonInstalar = [string]$cfgInicial.TextoBotonInstalar }
        if (-not $PSBoundParameters.ContainsKey('TextoBotonAplazar')) { $TextoBotonAplazar = [string]$cfgInicial.TextoBotonAplazar }
        if (-not $PSBoundParameters.ContainsKey('DiasLimite')) { $DiasLimite = [int]$cfgInicial.DiasLimite }
        if (-not $PSBoundParameters.ContainsKey('CuentaRegresivaMin')) { $CuentaRegresivaMin = [int]$cfgInicial.CuentaRegresivaMin }
        if (-not $PSBoundParameters.ContainsKey('CarpetaDestino')) { $CarpetaDestino = [string]$cfgInicial.CarpetaDestino }
        $datos = @{
            Instalador               = (Resolve-Path -LiteralPath $Instalador).ProviderPath
            Fabricante               = $Fabricante
            NombreApp                = $NombreApp
            Version                  = $Version
            Descripcion              = $Descripcion
            Arquitectura             = $Arquitectura
            ArgumentosInstalacion    = $ArgumentosInstalacion
            DesinstaladorExe         = $DesinstaladorExe
            ArgumentosDesinstalacion = $ArgumentosDesinstalacion
            NombreDeteccion          = $NombreDeteccion
            ProcesosACerrar          = $ProcesosACerrar
            PermitirAplazar          = (-not $SinAplazamiento)
            VecesAplazar             = $VecesAplazar
            MensajeFinal             = $MensajeFinal
            MensajeBienvenida        = $MensajeBienvenida
            TextoBotonInstalar       = $TextoBotonInstalar
            TextoBotonAplazar        = $TextoBotonAplazar
            DiasLimite               = $DiasLimite
            CuentaRegresivaMin       = $CuentaRegresivaMin
            SiemprePreguntar         = [bool]$SiemprePreguntar
            LogoApp                  = $LogoApp
            LogoAppEnVentanas        = [bool]$LogoAppEnVentanas
            CopiarCarpetaCompleta    = [bool]$CopiarCarpetaCompleta
            GenerarIntunewin         = [bool]$GenerarIntunewin
            CarpetaDestino           = $CarpetaDestino
            IncluirFuente            = [bool]$IncluirFuente
        }
        New-PaqueteCorporativo -Datos $datos | Out-Null
    }
    elseif ($ModoWeb) {
        Start-ModoWeb -Puerto $Puerto -SinNavegador:$SinNavegador
    }
    else {
        Show-Interfaz
    }
}
catch {
    Write-Registro -Mensaje $_.Exception.Message -Tipo Error
    exit 1
}
#endregion
