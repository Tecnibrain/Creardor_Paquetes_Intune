#Requires -Version 5.1
<#
.SYNOPSIS
    Fábrica de Paquetes PSADT para Intune (versión 2).

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

.EXAMPLE
    .\Nuevo-Paquete.ps1
    Abre la interfaz gráfica.

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
    [switch]$SinInterfaz
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
        default { $deteccion = "Script personalizado: Deteccion.ps1  (busca '$($D.NombreDeteccion)' versión >= $($D.Version))" }
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
    if ([string]::IsNullOrWhiteSpace($D.NombreDeteccion)) { $D.NombreDeteccion = $D.NombreApp }
    if ([string]::IsNullOrWhiteSpace($D.Descripcion)) {
        $D.Descripcion = "Instala $($D.NombreApp) $($D.Version) de $($D.Fabricante)."
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
    if ([string]::IsNullOrWhiteSpace($D.Version))   { $faltantes += 'Versión' }
    if ($D.Tipo -eq 'EXE' -and [string]::IsNullOrWhiteSpace($D.DesinstaladorExe)) { $faltantes += 'Desinstalador (obligatorio para EXE)' }
    if ($faltantes.Count -gt 0) { throw ('Faltan datos: ' + ($faltantes -join ', ')) }
    if ($D.Tipo -eq 'EXE' -and [string]::IsNullOrWhiteSpace($D.ArgumentosInstalacion)) {
        Write-Registro -Mensaje 'El EXE no tiene argumentos silenciosos. Si el instalador muestra ventanas, Intune quedará esperando.' -Tipo Aviso
    }
    if ($D.Tipo -eq 'PS1' -and -not $D.RutaRelativaDesinstalador) {
        Write-Registro -Mensaje 'PS1 sin script de desinstalación: al desinstalar solo se quitará la marca de registro.' -Tipo Aviso
    }

    $D.NombrePaquete = Get-NombreLimpio ('{0}_{1}_{2}_{3}' -f $D.Fabricante, $D.NombreApp, $D.Version, $D.Arquitectura)
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

    Write-Registro -Mensaje "Paquete listo en: $($D.RutaSalida)" -Tipo OK
    Write-Registro -Mensaje "Carpeta fuente (para editar): $($D.RutaPaquete)"
    Write-Registro -Mensaje 'Siguiente paso: probar como SYSTEM (ver FichaIntune.txt) antes de subir a Intune.'
    return @{ RutaPaquete = $D.RutaPaquete; RutaSalida = $D.RutaSalida; Intunewin = $intunewin }
}
#endregion

#region ===================== Interfaz gráfica =====================
$script:XamlInterfaz = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Fábrica de Paquetes PSADT para Intune" Width="1180" Height="820" MinWidth="960" MinHeight="600"
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
      <StackPanel>
        <Button x:Name="navCrear" Style="{StaticResource BotonMenu}" Content="📦   Crear paquete"/>
        <Button x:Name="navPlantilla" Style="{StaticResource BotonMenu}" Content="🎨   Plantilla corporativa"/>
        <Button x:Name="navAyuda" Style="{StaticResource BotonMenu}" Content="❓   Cómo funciona"/>
        <Separator Margin="0,14"/>
        <TextBlock Style="{StaticResource Ayuda}" Margin="6,0" Text="1. Arrastra el instalador.&#x0a;2. Revisa los datos.&#x0a;3. Mira la vista previa.&#x0a;4. Crea el paquete."/>
      </StackPanel>
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
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Style="{StaticResource Etiqueta}" Text="Versión"/>
                  <TextBox x:Name="txtVersion"/>
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
              <TextBlock Style="{StaticResource Ayuda}" Text="Si escribes algo, REEMPLAZA el texto estándar de PSADT. Vacío = texto estándar. Formato: [bold]negrita[/bold]  [italic]cursiva[/italic]  [accent]color[/accent]  [url]https://...[/url]"/>
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
        foreach ($par in @(@('Crear', $c.pagCrear, $c.navCrear), @('Plantilla', $c.pagPlantilla, $c.navPlantilla), @('Ayuda', $c.pagAyuda, $c.navAyuda))) {
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
    $c.navPlantilla.Add_Click({ & $irA 'Plantilla' })
    $c.navAyuda.Add_Click({ & $irA 'Ayuda' })

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
                $cfg = Get-ConfigCorporativa
                if ([string]::IsNullOrWhiteSpace($c.txtEmpresa.Text)) { throw 'Escribe el nombre de la empresa.' }
                $hex = $c.txtColor.Text.Trim()
                if (-not $hex.StartsWith('#')) { $hex = '#' + $hex }
                if ($hex -notmatch '^#[0-9A-Fa-f]{6}$') { throw 'El color debe tener el formato #RRGGBB (ej. #2C2A29).' }

                $cfg.Empresa            = $c.txtEmpresa.Text.Trim()
                $cfg.ColorAcento        = $hex.ToUpper()
                $cfg.EstiloDialogo      = if ($c.cmbEstilo.SelectedIndex -eq 1) { 'Classic' } else { 'Fluent' }
                $cfg.Idioma             = switch ($c.cmbIdioma.SelectedIndex) { 0 { 'es' } 1 { 'en' } default { 'auto' } }
                $cfg.RutaLogs           = $c.txtLogs.Text.Trim()
                $cfg.VecesAplazar       = [int]$c.sldAplazarDef.Value
                $cfg.DiasLimite         = [int]$c.sldDiasDef.Value
                $cfg.CuentaRegresivaMin = [int]$c.sldCuentaDef.Value
                $cfg.MensajeBienvenida  = $c.txtBienvenidaDef.Text.Trim()
                $cfg.MensajeFinal       = $c.txtMensajeDef.Text.Trim()
                $cfg.TextoBotonInstalar = $c.txtBotonInstalarDef.Text.Trim()
                $cfg.TextoBotonAplazar  = $c.txtBotonAplazarDef.Text.Trim()

                if ($null -ne $script:LogoPendiente) {
                    if ($script:LogoPendiente) {
                        # Se lee primero en memoria por si el archivo elegido es el mismo logo guardado antes.
                        $bytes = [System.IO.File]::ReadAllBytes($script:LogoPendiente)
                        $nombreLogo = 'LogoEmpresa' + [System.IO.Path]::GetExtension($script:LogoPendiente).ToLowerInvariant()
                        Get-ChildItem -LiteralPath $script:Recursos -Filter 'LogoEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
                        [System.IO.File]::WriteAllBytes((Join-Path $script:Recursos $nombreLogo), $bytes)
                        $cfg.Logo = $nombreLogo
                    }
                    else {
                        Get-ChildItem -LiteralPath $script:Recursos -Filter 'LogoEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
                        $cfg.Logo = ''
                    }
                }

                if ($null -ne $script:BannerPendiente) {
                    if ($script:BannerPendiente) {
                        $bytesBanner = [System.IO.File]::ReadAllBytes($script:BannerPendiente)
                        $nombreBanner = 'BannerEmpresa' + [System.IO.Path]::GetExtension($script:BannerPendiente).ToLowerInvariant()
                        Get-ChildItem -LiteralPath $script:Recursos -Filter 'BannerEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
                        [System.IO.File]::WriteAllBytes((Join-Path $script:Recursos $nombreBanner), $bytesBanner)
                        $cfg.Banner = $nombreBanner
                    }
                    else {
                        Get-ChildItem -LiteralPath $script:Recursos -Filter 'BannerEmpresa.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force
                        $cfg.Banner = ''
                    }
                }

                $sinImportarLogo = ($null -ne $script:LogoPendiente)
                $sinImportarBanner = ($null -ne $script:BannerPendiente)
                Save-ConfigCorporativa -Cfg $cfg
                Update-PlantillaCorporativa -Cfg $cfg -NoImportarLogo:$sinImportarLogo -NoImportarBanner:$sinImportarBanner
                & $cargarPlantilla
                Write-Registro -Mensaje 'Plantilla guardada: logo, banner y configuración actualizados.' -Tipo OK
                Write-Registro -Mensaje 'Los paquetes creados ANTES no cambian: vuelve a crearlos para que tomen la nueva plantilla.' -Tipo Aviso
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
    else {
        Show-Interfaz
    }
}
catch {
    Write-Registro -Mensaje $_.Exception.Message -Tipo Error
    exit 1
}
#endregion
