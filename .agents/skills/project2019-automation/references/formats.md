# Formatos de entrada y salida

Usa JSON UTF-8, no comandos incrustados ni datos separados por tabuladores enviados al teclado. Los valores numéricos son números JSON; no escribas `"2 días"` ni `"480"` en `durationMinutes`. Los scripts rechazan `null`, booleanos, cadenas en campos numéricos y campos de tarea no admitidos. Omitir un campo opcional conserva su valor; enviar `null` no equivale a omitirlo.

## Crear

```json
{
  "startDate": "2026-10-12T08:00:00",
  "tasks": [
    {"key": "inicio", "name": "Preparación", "durationMinutes": 480},
    {"key": "trabajo", "name": "Ejecución", "durationMinutes": 960, "predecessorKeys": ["inicio"]},
    {"key": "fin", "name": "Entrega", "durationMinutes": 0, "predecessorKeys": ["trabajo"]}
  ]
}
```

| Campo | Significado |
| --- | --- |
| `startDate` | Opcional; fecha/hora local ISO sin conversión implícita a UTC. Si falta, Project aplica su fecha de inicio predeterminada |
| `tasks` | Lista no vacía. Su orden determina el orden inicial de las filas |
| `key` | Clave única dentro del plan, no es el ID ni el UniqueID de Project |
| `name` | Texto no vacío; admite Unicode |
| `durationMinutes` | Número finito mayor o igual a cero, en minutos de trabajo |
| `predecessorKeys` | Claves de otras tareas del mismo plan; sin ciclos, sin autorreferencias, sin duplicados |
| `percentComplete` | Opcional; entero de 0 a 100 |

Los vínculos creados son fin-comienzo y sin demora. Las tareas se crean con `Manual=false`. No uses estos campos como si fueran recursos, restricciones o fechas fijas. Una duración cero puede representar un hito; la visualización concreta la resuelve Project.

La aplicación puede mostrar `?` si su configuración marca una duración como estimada. El snapshot comprueba los minutos y lee el indicador `Estimated`, pero los planes no cambian ese indicador explícitamente. Si es parte del requisito, añade una operación explícita y verifícala antes de anunciarla como soportada.

## Editar una copia

```json
{
  "updates": [
    {"uniqueId": 1, "percentComplete": 25},
    {"uniqueId": 2, "name": "Ejecución ampliada", "durationMinutes": 1440}
  ]
}
```

`uniqueId` es un entero positivo obtenido del snapshot del mismo archivo; no lo infieras de un ejemplo. Cada UniqueID aparece una vez. Campos admitidos: `name`, `durationMinutes` y `percentComplete`. Los campos omitidos se conservan. `UpdateCopy` no agrega/elimina tareas ni cambia predecesoras, calendarios o recursos.

Modificar una duración o el avance puede recalcular fechas y valores derivados. No establezcas fechas manualmente para forzar que la captura se parezca a una expectativa; revisa el calendario y el modelo.

## Reporte de verificación

`Inspect` produce `Version` y `Project`, con las filas leídas. No afirma que se hayan guardado o reabierto datos.

`Create` y `UpdateCopy` producen, si terminan correctamente:

- `Result: "PASS"` y `RequestedValuesVerified: true`.
- `ReopenedVerified: true`: los campos capturados coinciden al reabrir.
- `OriginalSnapshotUnchanged: true`: coinciden los campos observados del documento activo anterior; no significa que se hayan auditado todas sus propiedades.
- `SourceFileUnchanged: true` al editar una copia, o `null` al crear.
- `SHA256`, ruta de salida y snapshot del resultado.

`VerifiedUtc` es una marca UTC; presenta al usuario las horas en su zona cuando sea necesario. `Project.Start` es una representación local de Project y las fechas de tareas se serializan en ISO. No mezcles fechas de calendario con minutos de trabajo.

La comparación persistente incluye ruta, inicio, `HoursPerDay`, su conversión `MinutesPerDayConversion`, cantidad de tareas y campos de las filas: ID, UniqueID, nombre, duración, predecesoras, UniqueID de predecesoras, avance, programación manual, inicio, fin, resumen, hito, tipo/fecha de restricción y duración estimada. Excluye `Name` del proyecto y `Saved`, que cambiaron como metadatos de sesión al reabrir en la prueba.
