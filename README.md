# MsProject2019 Agent Kit

**Un procedimiento reproducible para que Codex cree y edite cronogramas, obtenga tablas y guarde capturas reales de Microsoft Project 2019 en Windows.**

El paquete combina la automatización COM de Project, instrucciones para agentes y scripts verificados. Project calcula y guarda el archivo `.mpp`; Codex prepara los datos y comprueba el resultado. La captura de pantalla usa una vía nativa de Windows cuando el capturador del plugin Computer Use no funciona.

Repositorio: [Sigilo-dev/MsProject2019-Agent-Kit](https://github.com/Sigilo-dev/MsProject2019-Agent-Kit). El código no depende del nombre de la carpeta, de una unidad concreta ni del perfil de Windows del autor.

## Empieza aquí

1. Clona este repositorio o descarga y extrae su ZIP. Abre la carpeta como proyecto local de Codex en el mismo Windows donde está instalado Microsoft Project.
2. Abre Project con un documento en blanco o el proyecto que quieras inspeccionar. Deja una sola instancia de la aplicación y resuelve los diálogos pendientes.
3. Pide a Codex:

   > Lee AGENTS.md y usa la skill project2019-automation de esta carpeta. Comprueba el entorno. Crea un cronograma con estas tareas, duraciones y dependencias, verifica el archivo al reabrirlo y entrega una tabla y una captura PNG. Guarda los resultados en una carpeta nueva dentro de runs. No alteres mis archivos originales.

4. Proporciona las tareas, la fecha de inicio y las dependencias. Para un archivo existente, proporciona su ruta y los cambios concretos.

Para usar la skill explícitamente en un entorno que admita esa sintaxis: `$project2019-automation`. Si no aparece en el selector, indica la ruta `.agents/skills/project2019-automation/SKILL.md`; no hace falta editar configuraciones globales para seguir el procedimiento.

## Qué hace y qué no promete

| Operación | Implementación y alcance |
| --- | --- |
| Leer tareas | Snapshot con identidad, duración, fechas, avance, modo, predecesoras, resumen, hito y restricciones; conversión de horas por día del proyecto |
| Crear un cronograma | Tareas automáticas, duración en minutos y vínculos fin-comienzo sin demora |
| Editar un `.mpp` | Copia nueva; modifica nombre, duración y avance por UniqueID |
| Calcular fechas | Lo hace Microsoft Project usando el calendario y las opciones del archivo |
| Hacer tablas | Crea filas en Project; exporta la tabla de tareas a CSV y Markdown desde datos verificados |
| Guardar imágenes | PNG real de la ventana visible de Project, incluyendo su interfaz |
| Conservar el original | No lo sobrescribe; comprueba su hash al editar una copia |
| Verificar | Reabre el resultado y compara los campos cubiertos, además de los valores solicitados |
| Recursos, costes, líneas base, calendarios especiales, subproyectos o tablas nativas personalizadas | Requieren ampliar y validar el procedimiento; no se anuncian como cubiertos por los scripts actuales |

No hay una garantía universal de «100 % a la primera»: influyen la edición instalada, los calendarios, la sesión gráfica, los permisos y el contenido del archivo. El propósito es evitar redescubrir los problemas ya resueltos y detenerse con evidencia cuando aparece uno nuevo. [VALIDATION.md](VALIDATION.md) especifica qué se comprobó.

## Requisitos

- Windows con Microsoft Project 2019 **de escritorio** instalado y funcionando. No se incluye ni se instala Project.
- Ejecución local en la misma sesión de usuario que Project.
- **Windows PowerShell 5.1** (`powershell.exe`), modo **STA**. No usar `pwsh`/PowerShell 7 ni WSL como sustituto directo.
- Una sola instancia de `WINPROJ`, con un documento activo. Puede haber varios documentos dentro de esa instancia.
- Una carpeta de salida con permiso de escritura y nombres de archivo nuevos.
- Para capturas: escritorio desbloqueado, Project visible y en primer plano. Computer Use es útil para identificar y activar la ventana; también puede hacerlo el usuario.
- Las herramientas y políticas del entorno deben permitir las operaciones. El paquete no concede acceso adicional ni evita restricciones.

No necesita API key, macros VBA, complementos de Office ni paquetes de Python para operar Project. Python solo se utilizó durante la preparación para validar el formato de la skill. Los scripts no desactivan políticas de ejecución o seguridad; un bloqueo de la máquina debe resolverse mediante el procedimiento aprobado para ese entorno.

## Por qué hay README, AGENTS y skill

| Archivo | Lectores y función |
| --- | --- |
| `README.md` | Personas: instalación local, alcance, ejemplos y requisitos |
| `AGENTS.md` | Codex al trabajar en esta carpeta: reglas breves y entrada operativa |
| `.agents/skills/project2019-automation/SKILL.md` | Agentes: escoger y ejecutar el flujo correcto |
| `references/` dentro de la skill | Detalles bajo demanda: entradas, COM, tablas, capturas y errores |
| `scripts/` dentro de la skill | Implementación reutilizable; evita que cada IA escriba un puente nuevo |
| `examples/` dentro de la skill | JSON sintético para crear y editar |
| `runs/` | Resultados locales, nunca contenido de publicación por defecto |

La estructura sigue la documentación de [AGENTS.md](https://learn.chatgpt.com/docs/agent-configuration/agents-md) y [skills locales de Codex](https://learn.chatgpt.com/docs/build-skills). La skill contiene sus propios recursos para que pueda copiarse completa a otro proyecto; no copies solo `SKILL.md`.

## Prueba rápida con scripts

Abre PowerShell en la raíz del paquete. Los ejemplos siguientes crean sus propias rutas absolutas. Si ejecutas comandos desde herramientas de Codex, mantén el directorio de trabajo en esa raíz.

```powershell
$skillDir = Join-Path $PWD.Path '.agents\skills\project2019-automation'
$scripts = Join-Path $skillDir 'scripts'
$runDir = Join-Path $PWD.Path ('runs\demo-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $runDir -ErrorAction Stop | Out-Null

powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Test-ProjectEnvironment.ps1"
if ($LASTEXITCODE -ne 0) { throw 'Corregir el entorno antes de continuar.' }

powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-Codex.ps1" `
  -Action Create -PlanPath "$skillDir\examples\plan.json" `
  -OutputPath "$runDir\base.mpp" -ReportPath "$runDir\create.json"
if ($LASTEXITCODE -ne 0) { throw 'Revisar el error; no repetir una mutacion a ciegas.' }

powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-Codex.ps1" `
  -Action UpdateCopy -SourcePath "$runDir\base.mpp" `
  -PlanPath "$skillDir\examples\updates.json" `
  -OutputPath "$runDir\edited.mpp" -ReportPath "$runDir\update.json"
if ($LASTEXITCODE -ne 0) { throw 'Revisar el error de edicion.' }

powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Export-ProjectTable.ps1" `
  -ReportPath "$runDir\update.json" -OutputDirectory "$runDir\table"
if ($LASTEXITCODE -ne 0) { throw 'Revisar el error de exportacion.' }
```

La fecha del ejemplo es fija para reproducibilidad: **12 de octubre de 2026**, no «hoy». Modifícala en tu JSON cuando corresponda. La duración de un día no debe suponerse siempre igual a 480 minutos: depende de la jornada y el calendario que uses.

Para inspeccionar sin editar:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-Codex.ps1" `
  -Action Inspect -ReportPath "$runDir\inspect.json"
```

Si sabes qué documento debe estar activo, añade `-ExpectedActiveProjectPath` con el valor exacto de `ActiveProject` devuelto por el chequeo de entorno. Un proyecto sin guardar puede llamarse simplemente `Project1`.

## Cambios sobre un archivo existente

1. Usa `Project-View.ps1 -Action Open` con la ruta de la fuente y un diario nuevo. `Inspect` por sí solo lee el documento activo, aunque hayas mencionado otro archivo en el chat.
2. Ejecuta `Inspect -ExpectedActiveProjectPath` con la ruta absoluta de esa fuente y guarda el reporte. Identifica las tareas por `UniqueID`, no por número de fila o nombre. Revisa resumen, hito, restricciones y conversión de días.
3. Ejecuta `Project-View.ps1 -Action Restore` **antes de editar la copia**. Después prepara `updates.json` y ejecuta `UpdateCopy` con fuente, JSON, salida y reporte nuevos. La fuente permanece como respaldo. Los comandos completos están en [workflow.md](.agents/skills/project2019-automation/references/workflow.md#inspeccionar-una-fuente-y-editar-su-copia) y el esquema en [formats.md](.agents/skills/project2019-automation/references/formats.md).
4. Si la fuente tiene cambios pendientes, el script se detiene: una copia de disco no representa esos cambios. Conserva primero la versión requerida dentro de la autorización del usuario.
5. Exige un reporte `PASS`. No sustituyas el archivo original automáticamente por el resultado.

Para calendarios compartidos, subproyectos o vínculos externos, revisa [workflow.md](.agents/skills/project2019-automation/references/workflow.md) antes de operar. Copiar un `.mpp` no independiza por sí mismo sus dependencias externas.

## Capturas después de los cambios

El flujo completo está en [captures.md](.agents/skills/project2019-automation/references/captures.md). Incluye los comandos para abrir el resultado, identificar y activar su ventana, guardar el PNG, revisarlo y volver al documento anterior.

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-View.ps1" `
  -Action Open -ProjectPath "$runDir\edited.mpp" -SessionPath "$runDir\view.json"
```

Después, Codex activa la ventana correcta mediante Computer Use o el usuario pone Project en primer plano. El helper de captura debe ejecutarse sin que una consola tape Project:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Guardar-Captura-Project.ps1" `
  -OutputPath "$runDir\after.png" -ExpectedTitle 'edited.mpp'
```

Es una captura real de los píxeles visibles. Si hay otra ventana flotante encima, puede aparecer en la imagen. Guardar el archivo no demuestra que el resultado sea legible: la IA debe abrirlo y comprobar título, filas, columnas, fechas y ausencia de contenido ajeno.

Finalmente:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-View.ps1" `
  -Action Restore -SessionPath "$runDir\view.json"
```

`Restore` cierra solo la vista que abrió el helper y vuelve al documento anterior. Si el resultado ya estaba abierto, lo conserva. El diario no autoriza otras acciones; no se reutiliza tras reiniciar Project.

## Resolver fallos sin improvisar

Consulta [troubleshooting.md](.agents/skills/project2019-automation/references/troubleshooting.md). Los casos ya observados incluyen:

- `FrameArrived timed out`: fallo de captura de Computer Use; no prueba un daño de Project.
- `coordinate input geometry is unavailable`: el clic puede necesitar una geometría que no existe aunque se lea el árbol de controles.
- `$app.Version` ausente: usar los wrappers `InvokeMember`; no reinstalar Office por ese síntoma.
- `Saved=false` después de reabrir en lectura: comparar los datos persistidos cubiertos; ese indicador de sesión no basta para juzgar la integridad del archivo.
- Project sin foco: activar la ventana correcta y volver a capturar; no guardar a ciegas el escritorio de otra aplicación.

No ejecutes `Quit`, `FileCloseAll`, `CalculateAll`, `/regserver`, reparaciones, cambios de licencia ni modificaciones del registro como solución rutinaria. No habilites macros ni desactives controles de seguridad para usar este paquete.

## Qué debe entregar la IA

- Archivo `.mpp` nuevo y ruta de la fuente conservada, si corresponde.
- Reporte JSON con la verificación realizada.
- Tabla CSV/Markdown si se solicitó.
- Captura PNG revisada si se solicitó; indicar cualquier limitación visual.
- Resumen de cambios y de aquello que quedó pendiente, sin afirmar que todos los campos de Project fueron certificados.

## Distribución

Conserva `.agents` al copiar o comprimir, ya que algunos selectores ocultan carpetas con punto inicial. No incluyas `runs/`, capturas con cuentas visibles, rutas personales, cronogramas reales ni reportes de usuarios. No se distribuyen ejecutables o bibliotecas de Microsoft.

El propietario debe elegir la licencia del paquete antes de presentarlo como software con una licencia de reutilización concreta. La licencia de Microsoft Project sigue siendo independiente.
