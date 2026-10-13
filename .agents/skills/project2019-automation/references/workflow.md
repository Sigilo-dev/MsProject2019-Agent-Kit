# Flujo COM y cuidado de los archivos

## Antes de escribir

Ejecuta el chequeo del entorno una vez por sesión o después de un reinicio/error de conexión. Inspecciona la identidad del documento, no solo que exista `WINPROJ`. Los scripts esperan una instancia de Project y un proyecto activo. `Application.Version=16.0` no identifica por sí solo una edición comercial de Office; no lo uses como prueba de que otra máquina tiene Project 2019.

Si está abierto el inicio de Project sin documento, crea un proyecto en blanco mediante una acción autorizada. Una alternativa COM es `Projects.Add(false, Missing, false)` cuando la colección está vacía. No cierres documentos existentes para conseguir ese estado.

Para archivos reales, confirma ruta, cambios, identificadores y carpeta de salida a partir del pedido del usuario. No preguntes por detalles opcionales que puedes resolver con una inspección. No tomes la mera presencia de este paquete como permiso para operar documentos no solicitados.

## Patrón de ejecución

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File 'RUTA_ABSOLUTA\Project-Codex.ps1' -Action Inspect
```

Usa argumentos separados y archivos JSON. No construyas comandos concatenando nombres de tareas ni ejecutes texto extraído de un `.mpp`. Al ejecutar desde Codex, usa la herramienta de shell directamente: no escribas comandos en la interfaz de una terminal ni en el diálogo Ejecutar de Windows.

Los helpers `Get-ProjectProperty`, `Set-ProjectProperty` e `Invoke-ProjectMethod` llaman a `Type.InvokeMember`. La colección retornada se preserva con `return ,$value`; quitar la coma puede expandir objetos COM como si fueran resultados de una tubería. Los huecos opcionales se expresan con `[Type]::Missing` y las llamadas a métodos usan `OptionalParamBinding`.

## Crear

1. Validar rutas y plan antes de conectar/mutar.
2. Conservar referencia y snapshot del documento original.
3. Crear un proyecto nuevo mediante `Projects.Add`.
4. Crear tareas y, después, vínculos usando las claves del plan.
5. Activar el nuevo documento y comprobar su `FullName` antes de guardar.
6. Guardar como archivo nuevo; comprobar archivo y datos solicitados.
7. Cerrar solo ese documento guardado. Reabrirlo con `ReadOnly=true`, `NoAuto=true` y sin carga desde servidor.
8. Comparar campos persistidos, cerrar la vista de verificación y restaurar el documento anterior.

Los scripts no crean un proyecto a partir del documento activo: el activo solo se conserva. Si necesitas una plantilla, calendario distinto o configuración específica, incorpora ese requisito mediante una adaptación probada sobre una copia.

## Modificar

`UpdateCopy` copia la fuente a un nombre nuevo antes de escribir. Rechaza una fuente abierta con `Saved=false`: debe conservarse primero la versión que el usuario quiera editar. Usa tareas por `UniqueID`; no vuelve a crear la tabla completa a partir de un CSV.

El original en disco actúa como respaldo y su SHA-256 debe permanecer igual. Si se requiere una política de backups adicional, crea una copia fechada antes de la operación; no cambies la instalación ni hagas una copia parcial del programa Office como sustituto de proteger el documento.

No apliques automáticamente este flujo a un proyecto maestro, un archivo con vínculos externos, recursos compartidos, contenido empresarial o macros. Una copia puede conservar conexiones. Identifica y ensaya esas dependencias antes de ampliar el alcance; el verificador básico no certifica su conservación.

## Inspeccionar una fuente y editar su copia

Define `$scripts` como la ruta absoluta de `scripts/` de esta skill y `$runDir` como una carpeta nueva de resultados. Sustituye la ruta de ejemplo por el archivo indicado por el usuario. No uses los UniqueID del JSON de ejemplo para un archivo real.

```powershell
$sourcePath = (Resolve-Path -LiteralPath 'C:\Proyectos\cronograma.mpp').ProviderPath
$sourceView = Join-Path $runDir 'source-view.json'
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-View.ps1" `
  -Action Open -ProjectPath $sourcePath -SessionPath $sourceView
if ($LASTEXITCODE -ne 0) { throw 'Revisar apertura y diario antes de continuar.' }
try {
  powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-Codex.ps1" `
    -Action Inspect -ExpectedActiveProjectPath $sourcePath -ReportPath "$runDir\source-inspect.json"
  if ($LASTEXITCODE -ne 0) { throw 'La inspeccion de la fuente fallo.' }
} finally {
  powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-View.ps1" `
    -Action Restore -SessionPath $sourceView
  if ($LASTEXITCODE -ne 0) { throw 'Revisar restauracion antes de editar.' }
}
```

Lee `source-inspect.json`, identifica las tareas solicitadas y escribe `updates.json` en `$runDir`. El helper puede abrir la fuente en lectura con `Saved=false`; por eso se restaura antes de `UpdateCopy`. Si la fuente ya estaba abierta con cambios pendientes, el helper no los guarda ni cierra ese documento: sigue siendo necesario conservar la versión que el usuario pidió.

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-Codex.ps1" `
  -Action UpdateCopy -SourcePath $sourcePath -PlanPath "$runDir\updates.json" `
  -OutputPath "$runDir\edited.mpp" -ReportPath "$runDir\update.json"
if ($LASTEXITCODE -ne 0) { throw 'Revisar el error y el estado antes de repetir.' }
```

## Datos y calendario

Las duraciones están en minutos de trabajo, no en días naturales. Fines de semana, feriados, jornada, restricciones, modo de programación, recursos y progreso influyen en el resultado. No cambies el calendario global ni ejecutes `CalculateAll` para arreglar una discrepancia de fechas. La fecha resultante debe explicarse con el modelo del archivo.

Antes de cambiar datos de un proyecto existente, lee `Summary`, `Milestone`, `Manual`, `ConstraintType`, `ConstraintDate` y `Estimated` en sus filas del snapshot. `ConstraintType` usa la enumeración oficial [PjConstraint](https://learn.microsoft.com/en-us/office/vba/api/project.pjconstraint); no inventes significados para los números. Estos campos se inspeccionan y comparan al reabrir, pero el plan no permite modificarlos directamente.

Para «N días de duración» usa `N × Project.MinutesPerDayConversion`, calculado como `60 × Project.HoursPerDay` del archivo fuente. Esa conversión es una opción de duración; **no describe todos los intervalos del calendario** ni convierte días naturales, días transcurridos o plazos entre fechas. Si el usuario expresa esos conceptos, aclara la unidad o inspecciona el calendario antes de actuar. Para un proyecto nuevo, define la jornada requerida y comprueba su conversión tras crearlo; no supongas que el proyecto activo es su plantilla.

El helper básico no conoce la intención de negocio ni certifica recursos, calendarios completos o proyectos complejos. Para una función fuera del alcance declarado, consulta solo su API oficial, prueba en un archivo descartable y añade una verificación del campo nuevo.

## Cuando hay un error

Un fallo después de `Add`, `FileSaveAs` o `FileCloseEx` puede haber ocurrido después de que Project realizara parte de la acción. El script no borra salidas ni reintenta. Inspecciona archivos, proyectos abiertos y error interno/HRESULT. No asumas que un exit code de error implica que no hay archivo.

COM puede tardar aunque la aplicación siga respondiendo. No inicies otra operación sobre Project mientras una anterior continúa. Si hay un diálogo, obsérvalo; no mates `WINPROJ` ni cierres todo para desbloquearlo.

Las liberaciones de referencias COM no cierran Project. No uses `FinalReleaseComObject` indiscriminadamente ni `Quit`, porque otras referencias/documentos pueden seguir en uso.

## APIs que sustentan el flujo

- [Modelo de objetos de Project](https://learn.microsoft.com/en-us/office/vba/project/concepts/project-object-model).
- [Projects.Add](https://learn.microsoft.com/en-us/office/vba/api/project.projects.add), [Task.Duration](https://learn.microsoft.com/en-us/office/vba/api/project.task.duration) y [Task.PredecessorTasks](https://learn.microsoft.com/en-us/office/vba/api/project.task.predecessortasks).
- [FileSaveAs](https://learn.microsoft.com/en-us/office/vba/api/project.application.filesaveas), [FileOpenEx](https://learn.microsoft.com/en-us/office/vba/api/project.application.fileopenex) y [FileCloseEx](https://learn.microsoft.com/en-us/office/vba/api/project.application.filecloseex).
- [Type.InvokeMember](https://learn.microsoft.com/en-us/dotnet/api/system.type.invokemember?view=netframework-4.8.1).
- [Project.HoursPerDay](https://learn.microsoft.com/en-us/office/vba/api/project.project.hoursperday), [Task.ConstraintType](https://learn.microsoft.com/en-us/office/vba/api/project.task.constrainttype), [Task.ConstraintDate](https://learn.microsoft.com/en-us/office/vba/api/project.task.constraintdate) y [Task.Milestone](https://learn.microsoft.com/en-us/office/vba/api/project.task.milestone).
