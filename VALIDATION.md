# Validación del paquete

Fecha: **13 de octubre de 2026**. Pruebas locales sobre archivos sintéticos descartables. No se modificó la instalación de Microsoft Project ni su registro, licencia o configuración de seguridad.

## Entorno observado

- Windows, sesión gráfica interactiva, una instancia de Microsoft Project Professional de escritorio.
- Windows PowerShell `5.1.19041.6456`, proceso de 64 bits, apartamento STA.
- Automatización `MSProject.Application`; `Application.Version = 16.0`. Ese número **no identifica por sí solo** Project 2019 frente a otras ediciones basadas en Office 16.
- Proyecto de prueba con conversión de 8 horas por día, tres tareas automáticas y vínculos fin-comienzo sin demora.

El usuario identificó su instalación como Project 2019. Esto es una validación en ese entorno, no una matriz de compatibilidad de todas las ediciones o equipos.

## Exportación nativa y flujo rápido (ampliación)

Se probó `Test-NativeImages.ps1` con **120 tareas**: diez cadenas fin-comienzo, un hito final, progreso del 25 %, texto Unicode y caracteres `&`/`<` escapados en XML. Project importó el plan, calculó, guardó y reabrió el MPP; los valores pedidos y el snapshot persistido coincidieron. Se inspeccionaron visualmente las tres imágenes siguientes, incluida la última fila.

| Operación | Dimensiones / alcance | Tiempo local |
| --- | --- | --- |
| Crear por XML, guardar y verificar al reabrir | 120 tareas | 9,35 s |
| Tabla completa → GIF → PNG | 747 × 2180 px; tareas 1–120 | 8,04 s |
| Tabla más Gantt → GIF → PNG | 1492 × 3339 px; tareas 1–120 | 9,05 s |
| Selección no contigua, columnas y título elegidos | 540 × 74 px; UniqueID 1, 60 y 120 | 11,04 s |

En cada exportación se verificaron los píxeles GIF/PNG, el hash intacto de la fuente y la restauración del documento previo. No hubo scroll ni capturas de pantalla. El test completo, incluida la creación y las tres exportaciones, tardó **38,15 s**, sin contar investigación, desarrollo ni revisión humana/visual. Son resultados de ese equipo y fixture; no una promesa para archivos reales arbitrarios.

La lectura por propiedades COM se midió alrededor de 350–400 ms por propiedad de tarea en una prueba puntual. Una exportación nativa del documento a XML tardó unos 272 ms y conservó nombre y estado `Saved`. Por eso el snapshot ahora usa exportación XML en bloque. No elimina la reapertura ni las comprobaciones.

También pasó la edición de una copia del archivo de 120 tareas con reapertura y conservación del hash fuente, y volvieron a pasar los nueve casos de rechazo de entradas. Las imágenes de datos no contienen la barra de cuenta de Office. Permanecen en `runs/` junto con los demás resultados locales.

La alternativa `Create -CreationMode Com` y la exportación CSV/Markdown siguieron funcionando con el snapshot en bloque. `Test-ImageConversion.ps1` comprobó preservación de píxeles, rechazo de GIF con dos frames (incluso sin demora), formato inválido y protección contra sobrescritura. Los trece scripts pasaron el análisis de sintaxis; el validador de skill también pasó.

Para repetir, con Project abierto y un documento activo:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Test-NativeImages.ps1" `
  -RunDirectory "$runDir\native-images-test"
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Test-ImageConversion.ps1" `
  -RunDirectory "$runDir\conversion-test"
```

Después abre los PNG: el test no sustituye la revisión de legibilidad. La exportación de una selección verifica sus UniqueID; `KeepRange` conserva el rango, pero Project puede recortar si agota memoria. No se certifican tamaños ilimitados.

## Resultados iniciales (antes de la ampliación)

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

No se certifican otros equipos, PowerShell 7, escritorio bloqueado, sesiones remotas, tamaños arbitrarios por encima del fixture de 120 tareas, asignaciones y costes, líneas base, calendarios complejos, subproyectos, recursos compartidos, macros ni archivos empresariales. Se probó una tabla temporal de columnas elegidas para exportar imágenes; no se certifica guardar cualquier diseño de vista en un MPP. Leer el estado de un hito o una restricción no equivale a haber validado su edición en todos los casos.

Los reportes, archivos MPP y capturas locales se conservan bajo `runs/` para inspección del propietario y están excluidos por `.gitignore`. El ZIP para compartir incluye documentación, scripts y JSON sintéticos; excluye esas evidencias y cualquier archivo de Microsoft Project.
