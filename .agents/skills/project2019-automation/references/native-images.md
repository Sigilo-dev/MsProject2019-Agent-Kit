# Exportación nativa: tabla o tabla más Gantt en un PNG

**Esta es la primera opción cuando se pide una imagen del cronograma, aunque tenga más filas que la pantalla.** No necesita capturas por scroll, portapapeles, Computer Use, Python, Pillow ni ImageMagick. Project genera un GIF; .NET lo convierte a PNG sin redimensionar y comprueba que los píxeles decodificados coincidan.

Una captura GDI sigue siendo útil cuando el pedido exige demostrar la interfaz de Project, un diálogo o un estado de la aplicación. No es necesaria para exportar datos que Project puede renderizar directamente.

## Un comando

Con Project abierto con un documento activo, ejecuta desde la raíz del paquete (o define `$scripts` con su ruta absoluta):

```powershell
$scripts = Join-Path $PWD.Path '.agents\skills\project2019-automation\scripts'
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Export-ProjectImage.ps1" `
  -ProjectPath 'C:\Proyectos\cronograma.mpp' `
  -OutputDirectory 'C:\Proyectos\entrega\imagen-gantt' -Mode TableAndGantt
if ($LASTEXITCODE -ne 0) { throw 'Revisar error y salidas; no repetir a ciegas.' }
```

Para **solo la tabla**, cambia a `-Mode Table`. La carpeta de salida debe ser nueva. El script deja:

- `project.gif`: exportación original de Project.
- `project.png`: misma resolución y mismos píxeles decodificados, verificados mediante SHA-256.
- `export.json`: filas, UniqueID, columnas, fechas, dimensiones, tiempos y comprobaciones.
- `render-<identificador>.mpp`: copia de trabajo conservada para diagnóstico; sus cambios de vista no se guardan.

La fuente debe estar guardada. Si ya está abierta con `Saved=false`, resuelve la versión requerida; no ignores los cambios pendientes. No hace falta ejecutar `Project-View Open` antes: el exporter abre su propia copia, prepara la vista, exporta, cierra esa copia y restaura el documento anterior. Comprueba que el hash del archivo fuente no cambió.

## Elegir solo lo solicitado

`-OptionsPath` acepta JSON UTF-8. Ejemplo:

```json
{
  "taskUniqueIds": [1, 60, 120],
  "columns": [
    {"field": "ID", "width": 5},
    {"field": "Name", "width": 45, "title": "Tarea"},
    {"field": "Duration", "width": 12, "title": "Duración"},
    {"field": "PercentComplete", "width": 14, "title": "Avance"}
  ]
}
```

Obtén los UniqueID del snapshot del archivo correcto. Omitir `taskUniqueIds` significa todas las tareas, incluidas las que no caben en pantalla. El script elimina filtros y agrupaciones, expande resúmenes y ordena por ID **en la copia**; no hereda un filtro visual como si fuese la selección pedida. Comprueba el conjunto de UniqueID de una selección parcial antes de exportarla. No convierte los números de fila en UniqueID.

Campos admitidos para columnas: `ID`, `UniqueID`, `Name`, `Duration`, `Start`, `Finish`, `Predecessors`, `PercentComplete`, `ResourceNames`. Se admiten entre 1 y 12 columnas; `width` es el ancho numérico de columna de Project (entero entre 3 y 100). Aumenta `Name` para nombres largos y revisa el resultado. Los identificadores `field` son fijos, independientes del idioma de Office; los títulos predeterminados los da la instalación. Un `title` explícito permite rotularlos en otro idioma. La tabla temporal no modifica `Global.MPT` ni añade un diseño permanente a la fuente.

Solo en `TableAndGantt`, puedes añadir **ambas** fechas:

```json
{"fromDate":"2026-10-11T00:00:00","toDate":"2026-10-29T00:00:00"}
```

Sin fechas, el rango cubre desde el inicio más temprano hasta el fin más tardío de las tareas seleccionadas, con margen de un día a cada lado. El zoom se ajusta a la selección. Fechas explícitas recortan intencionalmente el período exportado; no afirman mostrar todo el cronograma. Las fechas deben existir para exportar el Gantt de tareas manuales.

## Revisar una sola vez antes de entregar

Abre `project.png` con la herramienta de imágenes locales. Comprueba encabezados y columnas solicitadas, primera y última tarea de `export.json`, filas intermedias, nombres largos y ambos extremos del Gantt. Visualízala al 100 % para juzgar legibilidad: un visor puede reducir una imagen alta y hacer que parezca borrosa sin cambiar el archivo.

`Result: EXPORTED`, `SelectionVerified: true` y `PixelIdentical: true` prueban la selección y la conversión; **no prueban por sí solos que todo el texto se lea ni que Project no haya recortado por falta de memoria**. Por eso el reporte conserva `VisualReviewRequired: true`.

Si el resultado está recortado o es ilegible, cambia columnas, ancho o intervalo y exporta a otra carpeta. Para casos que superen la capacidad de Project, divide en bloques de tareas o períodos mediante exportaciones nativas y declara las partes, o usa PDF. No vuelvas por defecto a decenas de capturas con scroll. No hay un límite universal de tareas probado: la prueba del kit cubre 120 filas y una selección no contigua de tres tareas.

## Calidad y conversión independiente

GIF es un formato indexado: Project puede reducir los colores al generarlo. PNG conserva lo que contiene ese GIF; no añade resolución, colores ni detalles perdidos previamente. El helper no usa JPEG, escalado, reinterpolación ni IA para reconstruir contenido. Los GIF animados se rechazan para no perder frames sin avisar. El presupuesto del conversor es de 100 megapíxeles.

Para un GIF exportado manualmente:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Convert-ProjectGifToPng.ps1" `
  -InputPath 'C:\Proyectos\tabla.gif' -OutputPath 'C:\Proyectos\tabla.png'
```

## Equivalente en Project 2019 y API

En la interfaz: prepara **Hoja de tareas** para tabla o **Diagrama de Gantt** para tabla más barras; selecciona las filas; ve a **Tarea → Copiar → Copiar imagen → A archivo de imagen GIF**. Elige **Filas seleccionadas**, no solo filas en pantalla. En Gantt, indica también el intervalo de fechas.

El código usa `ViewApplyEx` con vistas incorporadas 23/4, una tabla temporal de columnas explícitas y `EditCopyPicture(Object=false, ForPrinter=2, SelectedRows=true, FromDate, ToDate, FileName, ScaleOption=1)`. `KeepRange` conserva la selección sin forzar su reducción. Microsoft advierte que puede recortarse si falta memoria; no se usa `Truncate`, ni `Scale` que distorsiona la proporción. Las rutas entregadas a COM se convierten explícitamente a `[string]`, evitando objetos enriquecidos devueltos por PowerShell que Project puede rechazar.

Fuentes oficiales: [Copiar una imagen del plan](https://support.microsoft.com/es-es/project/capture-a-picture-of-your-project-plan), [EditCopyPicture](https://learn.microsoft.com/en-us/office/vba/api/project.application.editcopypicture), [opciones de escala](https://learn.microsoft.com/en-us/office/vba/api/project.pjcopypicturescaleoption), [ViewApplyEx](https://learn.microsoft.com/en-us/office/vba/api/project.application.viewapplyex), [SetSplitBar](https://learn.microsoft.com/en-us/office/vba/api/project.application.setsplitbar).
