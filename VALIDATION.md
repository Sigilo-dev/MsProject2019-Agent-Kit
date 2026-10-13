# Validación del paquete

Fecha: **13 de octubre de 2026**. Pruebas locales sobre archivos sintéticos descartables. No se modificó la instalación de Microsoft Project ni su registro, licencia o configuración de seguridad.

## Entorno observado

- Windows, sesión gráfica interactiva, una instancia de Microsoft Project Professional de escritorio.
- Windows PowerShell `5.1.19041.6456`, proceso de 64 bits, apartamento STA.
- Automatización `MSProject.Application`; `Application.Version = 16.0`. Ese número **no identifica por sí solo** Project 2019 frente a otras ediciones basadas en Office 16.
- Proyecto de prueba con conversión de 8 horas por día, tres tareas automáticas y vínculos fin-comienzo sin demora.

El usuario identificó su instalación como Project 2019. Esto es una validación en ese entorno, no una matriz de compatibilidad de todas las ediciones o equipos.

## Resultados

| Prueba | Resultado observado |
| --- | --- |
| Chequeo del entorno | `READY`; conexión a la instancia existente y lectura del documento activo |
| Crear | `PASS`; tres tareas de 480, 960 y 480 minutos, dos dependencias y fecha inicial fija |
| Editar una copia | `PASS`; avance de la primera tarea a 25 %, segunda renombrada y ampliada a 1440 minutos |
| Reabrir resultados | `PASS`; campos persistidos cubiertos coinciden tras cerrar y abrir en lectura |
| Conservar fuente | SHA-256 de la fuente sin cambios en `UpdateCopy` |
| Conservar documento anterior | Snapshot cubierto del documento anterior sin cambios |
| Inspeccionar archivo concreto | Reporte JSON guardado y coincidencia con `ExpectedActiveProjectPath` |
| Información de programación | Lectura de Summary, Milestone, Manual, ConstraintType, ConstraintDate, Estimated, HoursPerDay y conversión a minutos |
| Exportar tablas | CSV y Markdown con las tres tareas; duración de 1440 minutos comprobada |
| Abrir para visualizar | Apertura en lectura, identidad comprobada y diario de propiedad de la apertura |
| Captura nativa | PNG real de 1928 × 1048 píxeles, inspeccionado visualmente; título, tres tareas, duración modificada, predecesoras y Gantt visibles |
| Restaurar vista | `RESTORED`; cierre de la vista abierta por el helper, hash intacto y regreso al documento previo |
| Entradas inválidas | Nueve casos rechazados antes de crear un proyecto, sin archivo `.mpp` de salida |
| Sintaxis y skill | Siete scripts sin errores de sintaxis; validador de skill aprobado |

Los nueve casos de entrada son: reporte que colisiona con la salida MPP, ciclo de dependencias, porcentaje fraccionario, campo desconocido, duración nula, duración booleana, duración como cadena, porcentaje nulo y porcentaje booleano. El primer intento del test de valores nulos detectó un problema del comparador de mensajes por el ajuste de líneas de PowerShell; se corrigió el comparador y los nueve casos pasaron. El rechazo del valor nulo funcionaba antes de esa corrección del test.

La imagen de prueba no muestra una columna numérica de porcentaje; ese dato se verificó mediante JSON y tabla. Algunas fechas quedan recortadas por el ancho de sus columnas. La imagen contiene información de cuenta en la interfaz y permanece en `runs/`; no se incluye en el paquete para compartir.

## Qué resolvió el diagnóstico

En la sesión de diagnóstico, Computer Use enumeraba y activaba ventanas y leía accesibilidad, pero su capturador devolvía `FrameArrived timed out` / `window capture timed out`. El clic por índice también devolvía `coordinate input geometry is unavailable`. Se observaron fallos de captura en más de una aplicación; no se atribuyó el problema exclusivamente a Project.

Se validó una alternativa: COM mediante reflexión para datos, Computer Use para seleccionar/activar ventanas cuando esté disponible y GDI para guardar los píxeles de Project en primer plano. No se afirma haber reparado internamente el capturador WGC del plugin.

## Repetir las pruebas

1. Ejecuta la prueba rápida del README en una carpeta nueva de `runs/`.
2. Revisa los reportes `create.json` y `update.json`; exige `PASS` y los campos de verificación correspondientes.
3. Ejecuta los casos de rechazo, que no necesitan una conexión a Project porque terminan antes de ella:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Test-PreflightGuards.ps1" `
  -ScratchPath "$runDir\guards"
if ($LASTEXITCODE -ne 0) { throw 'Revisar la prueba fallida.' }
```

4. Sigue la referencia de capturas, abre el PNG y restaura la vista.
5. Si amplías campos u operaciones, añade pruebas sobre archivos descartables para esa ampliación.

## Límites de la evidencia

No se certifican otros equipos, PowerShell 7, escritorio bloqueado, sesiones remotas, proyectos grandes, asignaciones y costes, líneas base, calendarios complejos, subproyectos, recursos compartidos, macros, archivos empresariales ni tablas/vistas nativas personalizadas. Leer el estado de un hito o una restricción no equivale a haber validado su edición en todos los casos.

Los reportes, archivos MPP y capturas locales se conservan bajo `runs/` para inspección del propietario y están excluidos por `.gitignore`. El ZIP para compartir incluye documentación, scripts y JSON sintéticos; excluye esas evidencias y cualquier archivo de Microsoft Project.
