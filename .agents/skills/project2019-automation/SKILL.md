---
name: project2019-automation
description: Crea y modifica cronogramas .mpp en Microsoft Project 2019 de escritorio, exporta tablas y guarda capturas reales en Windows. Usa COM y copias verificadas; para esta aplicación, no para Project web, Planner ni ProjectLibre.
---

# Microsoft Project 2019

Automatiza datos con los scripts COM de esta carpeta. Usa la aplicación instalada para calcular y guardar. Usa Computer Use para ubicar y activar ventanas; guarda capturas PNG mediante el helper GDI cuando la sesión y sus permisos lo permitan. No asumas que un fallo del capturador implica una instalación de Project dañada.

## Ruta corta

1. Localiza esta carpeta y crea una carpeta de resultados absoluta y nueva. Los scripts están en `scripts/`; no reconstruyas sus rutas desde el nombre de usuario de otra máquina.
2. Ejecuta `Test-ProjectEnvironment.ps1` con `powershell.exe -NoProfile -NonInteractive -STA -File`. Debe haber una sola instancia de Project y un documento activo. Si no está abierto, puedes abrir la aplicación instalada dentro de la autorización del usuario; no cierres otras instancias ni descartes cambios para satisfacer el chequeo.
3. Elige la operación en la tabla inferior. Para modificar un archivo existente, usa la secuencia `Project-View Open` de la fuente → `Inspect -ExpectedActiveProjectPath` → identifica UniqueID → `Project-View Restore` → `UpdateCopy`. La receta completa está en [workflow.md](references/workflow.md). Para uno nuevo, transforma los requisitos en el JSON de ejemplo.
4. Espera a que termine la operación. Exige `Result: PASS`, valores solicitados verificados y reapertura comprobada. Si falla, lee el estado antes de volver a escribir; no lances el mismo comando en bucle.
5. Para entregar tablas, exporta el reporte verificado. Para entregar imágenes, abre el resultado, activa su ventana, captura y revisa el PNG. Restaura el documento previo con el diario de vista.
6. Entrega rutas, cambios, evidencia y límites reales. No declares «100 % compatible» a partir de una prueba de tres tareas.

| Necesidad | Script o lectura |
| --- | --- |
| Inspeccionar el documento activo | `Project-Codex.ps1 -Action Inspect -ReportPath ...` |
| Crear tareas y vínculos fin-comienzo | `Project-Codex.ps1 -Action Create` + [formato de entradas](references/formats.md) |
| Editar nombre, duración o avance en una copia | `Project-Codex.ps1 -Action UpdateCopy` + [flujo COM](references/workflow.md) |
| Obtener tabla CSV y Markdown | `Export-ProjectTable.ps1` + [tablas](references/tables.md) |
| Abrir resultado para verlo y volver al anterior | `Project-View.ps1 -Action Open/Restore` |
| Guardar captura real | [capturas](references/captures.md), luego `Guardar-Captura-Project.ps1` |
| Resolver un error observado | [diagnóstico](references/troubleshooting.md), solo la fila relevante |

## Condiciones que no debes saltarte

- Ejecutor: **Windows PowerShell 5.1, STA**, misma sesión de usuario que Project. No ejecutes desde WSL ni asumas que los conectores web controlan el escritorio local.
- No sobrescribas salidas. `UpdateCopy` conserva el archivo fuente; cambios sin guardar en la fuente no están incluidos en una copia de disco.
- Activa y comprueba el documento de trabajo antes de cualquier `FileSave`, `FileSaveAs` o `FileCloseEx`: estos métodos actúan sobre el documento activo.
- No uses `Quit`, `FileCloseAll`, `CalculateAll`, `/regserver`, reparaciones de Office ni cambios globales para resolver el flujo normal.
- Los nombres de las tareas no son identificadores estables. Usa `UniqueID` del proyecto correcto; `ID` es una posición que puede cambiar.
- El acceso mediante `InvokeMember` de `ProjectCom.ps1` es intencional: el acceso directo `$app.Version` falló en la máquina de validación aunque COM funcionaba.
- La lectura de accesibilidad sin captura puede funcionar y, aun así, el clic fallar por falta de geometría. No improvises coordenadas ni reutilices handles de otra sesión.
- No eludas bloqueos de herramientas, permisos, escritorio bloqueado ni una detención del usuario usando el helper alternativo. Usa una captura solo dentro del acceso autorizado.
- Sin mutaciones concurrentes. La comparación posterior verifica campos concretos, no todos los recursos, calendarios, macros o vínculos externos del archivo.

Para el primer uso o archivos con calendario/recursos complejos, lee [workflow.md](references/workflow.md). Para operaciones rutinarias ya validadas, ejecuta directamente el script pertinente; evita releer documentación o investigar APIs que no necesitas.
