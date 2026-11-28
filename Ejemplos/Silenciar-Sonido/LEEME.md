# Desactivar el sonido de los PCs (Intune)

Script listo para la Fábrica de Paquetes (tipo **PS1**).

| Archivo | Qué hace |
|---|---|
| `Silenciar-Sonido.ps1` | Silencia todas las salidas de audio y deja el volumen en 0. Con `-Bloquear` además deshabilita el servicio *Audio de Windows* (el usuario no puede reactivarlo). |
| `Desinstalar-Silenciar-Sonido.ps1` | Revierte: rehabilita el servicio, quita el silencio y deja el volumen en 50 % (`-Volumen 30` para otro valor). |

## Crear el paquete

1. Abre `Crear-Paquete.cmd` y arrastra **`Silenciar-Sonido.ps1`**. El script de desinstalación se detecta solo.
2. **Argumentos de instalación:**
   - vacío → solo silenciar (el usuario puede volver a subir el volumen);
   - `-Bloquear` → sin sonido y sin poder activarlo.
3. Crea el paquete y publícalo como siempre (asignado como **Requerido** al grupo de equipos).

## A tener en cuenta

- Intune lo ejecuta **una vez**: si el usuario vuelve a subir el volumen, Intune no lo vuelve a silenciar (la detección es la marca de registro). Si debe ser permanente, usa `-Bloquear`.
- Con `-Bloquear` tampoco funcionan micrófono ni audio en Teams/llamadas, porque se apaga todo el servicio de audio.
- Prueba manual (como SYSTEM): `psexec -s -i powershell -ExecutionPolicy Bypass -File .\Silenciar-Sonido.ps1`
