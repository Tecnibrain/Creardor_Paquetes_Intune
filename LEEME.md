# Fábrica de Paquetes PSADT para Intune (v4.0)

Crea paquetes de **PSAppDeployToolkit 4.x** listos para **Intune** a partir de un **MSI, EXE o script PS1**, sin escribir código.

## Cómo se usa

1. Copia la carpeta `FabricaPaquetes` a cualquier disco del equipo de empaquetamiento, por ejemplo `D:\FabricaPaquetes`.
2. Haz doble clic en **`Crear-Paquete.cmd`**.
3. **Arrastra el instalador** (.msi, .exe o .ps1) a la zona punteada. Los datos se llenan solos.
4. Escribe la **Descripción para Company Portal** (tarjeta *2 · Información de la app*).
5. Ajusta las ventanas en la tarjeta *4* y mira la **vista previa en vivo** a la derecha.
6. Pulsa **Crear paquete**.

**Dónde queda el paquete:** en el panel derecho, *Dónde queda el paquete listo* → **Cambiar carpeta...** (puede ser una carpeta de red, por ejemplo `\\servidor\Paquetes\Intune`). La herramienta recuerda la carpeta. Ahí se crea `<Paquete>\` con el `.intunewin`, la ficha, la detección, la descripción y el logo. Si marcas *Incluir también la carpeta fuente*, se copia además la carpeta editable del paquete en `<Paquete>\Fuente`.

La interfaz tiene cuatro secciones en el menú de la izquierda:
- **Crear paquete:** el día a día.
- **Paquetes creados:** lista de todo lo que se ha creado con la Fábrica (fecha, quién, si tiene `.intunewin`), con botones para abrir la salida o la carpeta fuente. Tiene buscador.
- **Plantilla corporativa:** empresa, color, logo, idioma y valores por defecto.
- **Cómo funciona:** guía rápida dentro de la herramienta.

Ayudas mientras llenas el formulario:
- Los campos obligatorios que faltan se marcan **en rojo** (y en naranja los avisos, por ejemplo un EXE sin argumentos silenciosos).
- Encima del botón **Crear paquete** hay un resumen: qué falta o «Todo listo».
- En el menú de la izquierda, **Tu avance** marca con ✔ cada paso completo.
- Atajos: **Ctrl+O** elegir instalador, **Ctrl+Enter** crear paquete, **F5** ver ventanas reales.

Abajo, la barra de estado muestra lo que está pasando. Haz clic en ella para ver el registro completo.

## Versión web

La misma Fábrica, en el navegador. Ábrela con **`Crear-Paquete-Web.cmd`** o con el botón **🌐 Abrir versión web** del menú.

- Se abre Edge en una ventana propia. La ventana negra (consola) **es la Fábrica**: no la cierres mientras uses la página. Para terminar, pulsa **Cerrar Fábrica** en la página (o Ctrl+C en la consola).
- **Solo funciona en este equipo:** escucha en `127.0.0.1` (no acepta conexiones de otros equipos) y usa un código secreto que cambia cada vez que se abre. No necesita IIS, permisos de administrador ni Internet.
- Tiene lo mismo que la de escritorio: crear paquete, vista previa, ver ventanas reales, paquetes creados y plantilla corporativa. Además tiene tema oscuro (botón **◐ Tema**).
- **El instalador**, de tres formas:
  - **Clic en la zona punteada:** abre el buscador de archivos de Windows (puede quedar detrás del navegador: búscalo en la barra de tareas).
  - **Pegar la ruta** completa, por ejemplo `D:\Fuentes\MiApp\setup.msi`, y pulsar **Leer**.
  - **Arrastrar el archivo:** el navegador no entrega la ruta, así que **solo se copia ese archivo**. Si el instalador trae archivos adicionales o es un `.ps1` con su script de desinstalación, usa el clic o la ruta.
- Mientras se crea el paquete la página muestra «Creando el paquete…»; el detalle en vivo se ve en la consola.

Si Edge no está instalado, se abre el navegador por defecto. Para usar un puerto fijo: `Nuevo-Paquete.ps1 -ModoWeb -Puerto 8765`.

Listo. En `Salida\<Paquete>\` quedan:

| Archivo | Para qué |
|---|---|
| `<Paquete>.intunewin` | Se sube a Intune |
| `FichaIntune.txt` | Todo lo que hay que poner en Intune (comandos, códigos, detección) |
| `Deteccion.ps1` | Regla de detección |
| `Logo_Intune.png` | Ícono de la app para Company Portal |
| `Descripcion.txt` | Descripción para pegar en Intune |

## Primera vez (automático)

La primera vez que se abre, la herramienta hace esto sola:

- **Busca el módulo PSAppDeployToolkit** en este orden: la carpeta `Herramientas`, luego el módulo instalado en el equipo y, por último, lo descarga de la PowerShell Gallery si hay Internet. Si lo descarga, queda guardado en `Herramientas` para las siguientes veces y para quien reciba la carpeta.
- **Crea la plantilla corporativa** en `Recursos` con la empresa, el color, el idioma español, el logo y el banner generados.

> Si el equipo no tiene Internet y no hay módulo, la herramienta lo avisa. En un equipo con Internet ejecuta
> `Save-Module -Name PSAppDeployToolkit -Path C:\Temp\PSADT` y copia `C:\Temp\PSADT\PSAppDeployToolkit` dentro de `Herramientas`.
> **Consejo:** ábrela una vez en un equipo con Internet antes de compartir la carpeta. Así ya viaja con todo.

`IntuneWinAppUtil.exe` (oficial de Microsoft) ya viene incluido en `Herramientas`.

## Pestaña "Plantilla corporativa"

Se configura una vez y aplica a **todos** los paquetes nuevos:

- **Empresa:** aparece en el título de las ventanas y en el banner.
- **Color corporativo:** color de botones y fondo del banner.
- **Logo:** eliges un PNG. Si no eliges ninguno, se generan las iniciales de la empresa.
- **Banner (estilo Classic):** **Subir banner...** (PNG de 900 × 125 px) o **Generar automático**.
- Si reemplazas a mano una imagen en `Recursos\Assets`, la herramienta lo detecta al guardar y **la conserva**: no la sobrescribe.
- **Estilo:** Fluent (moderno) o Classic.
- **Idioma:** español, inglés o automático.
- **Ruta de logs**, **aplazamientos por defecto** y **mensaje final por defecto**.

Pulsa **Guardar plantilla** y verás la vista previa.

## ¿Por qué a veces no salen las ventanas? (importante)

PSADT decide solo si muestra ventanas (modo **Auto**):

- **Hay procesos a cerrar y la app está ABIERTA** → muestra la ventana completa: lista de apps con su ícono, cuenta regresiva, aplazamientos y fecha límite.
- **Hay procesos a cerrar y la app está CERRADA** → instala **en silencio**, porque no hay nada que molestar. Es el comportamiento oficial.
- **No hay nadie conectado** → siempre en silencio.

Para probar cómo se ven las ventanas usa el botón **Vista previa de ventanas**. Muestra las ventanas reales de PSADT con tu configuración y **no instala nada**. Si no pusiste procesos a cerrar, abre el Bloc de notas para que veas la lista de apps.

Si quieres que salgan aunque la app esté cerrada, marca **"Mostrar ventanas aunque la app esté cerrada"**.

Opciones de la ventana de bienvenida:
- **Procesos a cerrar:** salen listados con su ícono. Puedes darles nombre visible: `excel=Microsoft Excel`.
- **Aplazar + veces:** muestra "Aplazamientos restantes".
- **Fecha límite (días):** muestra la fecha hasta la que puede aplazar.
- **Cuenta regresiva (minutos):** al llegar a 0 cierra las apps e instala.
- **Texto de los botones:** cambia «Instalar» y «Aplazar» por lo que quieras (vacío = texto estándar).
- **Mensaje en la bienvenida:** admite `[bold]negrita[/bold]`, `[italic]cursiva[/italic]`, `[accent]color[/accent]` y `[url]https://...[/url]`.

## El logo

- **Logo de la empresa** (pestaña *Plantilla corporativa*): eliges la imagen y pulsas **Guardar**. **Solo afecta a los paquetes que crees después.** Los que ya existían hay que volver a crearlos.
- **Logo de la app** (pestaña *Crear paquete*): si es un EXE se extrae su ícono solo, o eliges una imagen. Se usa para `Logo_Intune.png` (Company Portal). Si marcas **"Mostrar este logo en las ventanas"**, también aparece en las ventanas de ese paquete.

## Qué llenar según el tipo

| | MSI | EXE | PS1 |
|---|---|---|---|
| Datos (nombre, versión, fabricante) | Automático | Automático | Nombre automático; versión `1.0.0` (cámbiala) |
| Argumentos | Opcional (`PROPIEDAD=valor`) | **Silenciosos** (`/S`, `/quiet`…) | Opcional |
| Desinstalador | No aplica | **Obligatorio**: ruta en el equipo (`$env:ProgramFiles\App\uninstall.exe`) | Opcional: tu `Desinstalar.ps1`. Se detecta solo si se llama *desinstalar/uninstall* |
| Detección | Código de producto MSI | Nombre en Programas y características | Marca de registro que crea el paquete |

**Scripts PS1:**
- Se copia **toda la carpeta** del script.
- Tu script debe terminar con **`exit 0`** si todo salió bien.
- El paquete deja la marca `HKLM\SOFTWARE\<Empresa>\Paquetes\<App>` con la versión. Intune la usa para saber si está instalado y la borra al desinstalar.

## Probar antes de publicar (obligatorio)

`FichaIntune.txt` trae los comandos exactos:

```cmd
psexec -s -i cmd
cd /d D:\FabricaPaquetes\Paquetes\<Paquete>
Invoke-AppDeployToolkit.exe -DeploymentType Install
powershell -ExecutionPolicy Bypass -File D:\FabricaPaquetes\Salida\<Paquete>\Deteccion.ps1
Invoke-AppDeployToolkit.exe -DeploymentType Uninstall
```

## Publicar en Intune

1. **Aplicaciones → Windows → Agregar → Aplicación Windows (Win32)** → sube el `.intunewin`.
2. Copia los valores de `FichaIntune.txt` y sube `Logo_Intune.png` como logo.
3. Asigna como **Disponible** a un grupo piloto y luego pásala a **Requerido**.

## Estructura

```
FabricaPaquetes\
├── Crear-Paquete.cmd      ← doble clic (escritorio)
├── Crear-Paquete-Web.cmd  ← doble clic (versión web)
├── Nuevo-Paquete.ps1      ← la herramienta
├── LEEME.md
├── Herramientas\          ← IntuneWinAppUtil.exe (+ módulo PSADT)
├── Recursos\              ← plantilla corporativa (se crea sola) e Historial.json
├── Paquetes\              ← un folder por paquete
└── Salida\                ← lo que se sube a Intune
```

## Casos que requieren editar a mano `Invoke-AppDeployToolkit.ps1`

- Un MSI con transform (`.mst`) o varios instaladores.
- Copiar licencias o configuraciones después de instalar.
- Quitar primero una versión anterior distinta.
- Un instalador que **pide datos al usuario**.

## Seguridad

- La versión web escucha solo en `127.0.0.1`, valida las cabeceras `Host` y `Origin` y exige un código aleatorio por sesión en cada petición. Desde la página solo se pueden abrir en el Explorador las carpetas de la Fábrica y de sus paquetes. Los archivos arrastrados se guardan en `%TEMP%\FabricaWeb_*` y se borran al cerrar la Fábrica.
- `Crear-Paquete.cmd` usa `-ExecutionPolicy Bypass` solo en el equipo de empaquetamiento. Si la política lo exige, firma los scripts con el certificado interno.
- El módulo PSADT valida su propia firma digital: usa solo la versión oficial (PowerShell Gallery o GitHub de PSAppDeployToolkit).
- No pongas contraseñas ni tokens en los argumentos: quedan en los logs y en la caché de Intune.
- Todo se ejecuta como SYSTEM. Revisa el script generado antes de publicar.
- `IntuneWinAppUtil.exe` está sujeto a los términos de licencia de Microsoft (PDF incluido en `Herramientas`).
