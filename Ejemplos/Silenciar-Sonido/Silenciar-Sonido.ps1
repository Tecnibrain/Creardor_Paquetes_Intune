<#
.SYNOPSIS
    Desactiva el sonido del equipo (todas las salidas de audio).

.DESCRIPTION
    Pensado para empaquetarse con la Fábrica de Paquetes (tipo PS1) y desplegarse con Intune como SYSTEM.

    Sin parámetros: silencia (mute) todas las salidas de audio activas y deja el volumen en 0.
    El usuario puede volver a subir el volumen si quiere.

    Con -Bloquear: además detiene y deshabilita el servicio «Audio de Windows» (Audiosrv).
    El equipo queda sin sonido y el usuario no puede activarlo (necesita permisos de administrador).

    Para revertir: Desinstalar-Silenciar-Sonido.ps1 (la Fábrica lo detecta solo como script de desinstalación).

.PARAMETER Bloquear
    Deshabilita el servicio de audio para que el sonido no se pueda volver a activar.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Silenciar-Sonido.ps1
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Silenciar-Sonido.ps1 -Bloquear
#>
[CmdletBinding()]
param(
    [switch]$Bloquear
)

$ErrorActionPreference = 'Stop'

$codigo = @'
using System;
using System.Runtime.InteropServices;

namespace FabricaAudio
{
    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceEnumerator
    {
        void EnumAudioEndpoints(int dataFlow, int stateMask, out IMMDeviceCollection devices);
    }

    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceCollection
    {
        void GetCount(out uint count);
        void Item(uint index, out IMMDevice device);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDevice
    {
        void Activate(ref Guid iid, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object anInterface);
        void OpenPropertyStore(int access, out IntPtr properties);
        void GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
    }

    [ComImport, Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioEndpointVolume
    {
        void RegisterControlChangeNotify(IntPtr notify);
        void UnregisterControlChangeNotify(IntPtr notify);
        void GetChannelCount(out uint count);
        void SetMasterVolumeLevel(float levelDB, ref Guid context);
        void SetMasterVolumeLevelScalar(float level, ref Guid context);
        void GetMasterVolumeLevel(out float levelDB);
        void GetMasterVolumeLevelScalar(out float level);
        void SetChannelVolumeLevel(uint channel, float levelDB, ref Guid context);
        void SetChannelVolumeLevelScalar(uint channel, float level, ref Guid context);
        void GetChannelVolumeLevel(uint channel, out float levelDB);
        void GetChannelVolumeLevelScalar(uint channel, out float level);
        void SetMute([MarshalAs(UnmanagedType.Bool)] bool mute, ref Guid context);
        void GetMute([MarshalAs(UnmanagedType.Bool)] out bool mute);
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumeratorCom { }

    public static class Salidas
    {
        // Silencia (o activa) todas las salidas activas. Devuelve cuántas cambió.
        public static int Silenciar(bool silenciar, float volumen)
        {
            IMMDeviceEnumerator enumerador = (IMMDeviceEnumerator)(new MMDeviceEnumeratorCom());
            IMMDeviceCollection coleccion;
            enumerador.EnumAudioEndpoints(0 /* eRender */, 1 /* DEVICE_STATE_ACTIVE */, out coleccion);
            uint total;
            coleccion.GetCount(out total);
            Guid iid = typeof(IAudioEndpointVolume).GUID;
            Guid contexto = Guid.Empty;
            int cambiadas = 0;
            for (uint i = 0; i < total; i++)
            {
                IMMDevice dispositivo;
                coleccion.Item(i, out dispositivo);
                object o;
                dispositivo.Activate(ref iid, 23 /* CLSCTX_ALL */, IntPtr.Zero, out o);
                IAudioEndpointVolume vol = (IAudioEndpointVolume)o;
                if (volumen >= 0) { vol.SetMasterVolumeLevelScalar(volumen, ref contexto); }
                vol.SetMute(silenciar, ref contexto);
                cambiadas++;
                Marshal.ReleaseComObject(vol);
                Marshal.ReleaseComObject(dispositivo);
            }
            Marshal.ReleaseComObject(coleccion);
            Marshal.ReleaseComObject(enumerador);
            return cambiadas;
        }
    }
}
'@

try {
    if (-not ('FabricaAudio.Salidas' -as [type])) { Add-Type -TypeDefinition $codigo -Language CSharp }

    $servicio = Get-Service -Name 'Audiosrv' -ErrorAction SilentlyContinue
    if ($servicio -and $servicio.Status -ne 'Running') {
        Write-Output 'El servicio Audio de Windows no está en ejecución: no hay nada que silenciar.'
    }
    else {
        $n = [FabricaAudio.Salidas]::Silenciar($true, 0.0)
        Write-Output "Salidas de audio silenciadas: $n"
    }

    if ($Bloquear -and $servicio) {
        Stop-Service -Name 'Audiosrv' -Force -ErrorAction SilentlyContinue
        Set-Service -Name 'Audiosrv' -StartupType Disabled
        Write-Output 'Servicio Audio de Windows detenido y deshabilitado.'
    }

    exit 0
}
catch {
    Write-Output "Error al desactivar el sonido: $($_.Exception.Message)"
    exit 1
}
