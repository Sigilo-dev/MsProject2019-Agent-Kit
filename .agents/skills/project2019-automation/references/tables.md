# Tablas: tres necesidades distintas

## Filas de tareas dentro de Project

Para crear una tabla con tareas, duraciones y predecesoras, usa `Create`. El resultado es un `.mpp` nativo: las filas aparecen en la tabla de la vista Gantt. No hay que pegar bloques por teclado ni reconstruir el archivo desde una imagen.

`UpdateCopy` modifica las celdas admitidas por UniqueID. Si se piden inserciones, eliminaciones o jerarquías de tareas, esos casos necesitan una ampliación específica; no simules una eliminación vaciando nombres.

## Tabla para entregar o revisar

Exporta desde el JSON de `Inspect` o desde el reporte `PASS` de la última operación:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File 'RUTA\Export-ProjectTable.ps1' `
  -ReportPath 'RUTA\update.json' -OutputDirectory 'RUTA\table'
```

Produce `tasks.csv` y `tasks.md` con ID, UniqueID, tarea, minutos, inicio, fin, predecesoras y porcentaje. No modifica Project ni calcula fechas por su cuenta. Un reporte es un snapshot: si cambió el archivo después, genera uno nuevo antes de entregar.

El CSV usa UTF-8 con BOM, separador coma y comillas estándar. En Excel con separador regional distinto, importa indicando coma en lugar de alterar el archivo a ciegas. Los textos que empiezan como fórmulas se exportan como texto seguro; el `.mpp` no se modifica. Markdown escapa separadores de columna y saltos de línea.

Esta tabla no es un formato completo de intercambio de Project: no contiene calendarios, recursos, costes, líneas base ni todos los campos. No la uses para sobrescribir un `.mpp` existente suponiendo que conserva todo.

## Tabla nativa personalizada: columnas, vistas y formato

No está implementada por el exporter CSV. Si el pedido exige nuevas columnas, anchuras, tablas de costes/recursos o diseños nativos, trabaja en una copia y consulta los métodos oficiales de [TableEditEx](https://learn.microsoft.com/en-us/office/vba/api/project.application.tableeditex) y [TableApply](https://learn.microsoft.com/en-us/office/vba/api/project.application.tableapply). Usa nombres de tabla/columna obtenidos de la instalación, porque pueden estar traducidos.

Verifica la vista visualmente, guarda y reabre para comprobar que el formato solicitado se conservó. No edites `Global.MPT` ni presentes una imagen o Markdown como si fueran una tabla nativa guardada. Si amplías el código, añade prueba del comportamiento nuevo; no anuncies esa función como probada solo porque aparece en esta referencia.
