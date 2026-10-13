# Errores conocidos y respuesta mínima

| Síntoma | Interpretación y acción |
| --- | --- |
| `Expected one Project application` | No hay instancia o hay varias. Observa procesos/ventanas. Abre Project si está autorizado y ausente; no cierres trabajo ajeno |
| `GetActiveObject` no encuentra la instancia | Puede no estar registrada aún, no haber documento o pertenecer a otra sesión/elevación. Comprueba sesión y documento; no uses `/regserver` como primer paso |
| `ActiveProject` falla | Puede estar abierta solo la pantalla de inicio o un diálogo. Crea/abre el documento autorizado |
| `Version` o `Projects` ausente por acceso directo | Utiliza los wrappers de `ProjectCom.ps1`; la prueba observó COM funcional pese a ese comportamiento |
| Se ejecuta con `pwsh` | Repetir la operación solo si aún no hubo mutación, usando Windows PowerShell 5.1 STA |
| `FrameArrived timed out` / `window capture timed out` | El capturador no entregó un frame. No seguir reintentando en bucle; usar el procedimiento de capturas si está permitido |
| `coordinate input geometry is unavailable` | Poder leer el árbol no garantiza poder hacer clic. COM para datos; activación de ventana y captura GDI para evidencia |
| `Foreground window is not Microsoft Project` | No se guardó imagen. Activar el Project correcto y reintentar una captura con nombre nuevo |
| PNG negro, tapado o equivocado | No usarlo como prueba. Comprobar escritorio, foco, título y superposiciones. No falsificar ni reconstruir una captura |
| Cambia el foco al maximizar | Volver a observar/activar antes de capturar. No asumir que maximizar haya ensanchado la ventana |
| Fuente con `Saved=false` | Puede haber cambios pendientes o metadatos de sesión. No copiar a ciegas una versión de disco que omita cambios |
| `Saved=false` al reabrir el resultado en lectura | Fue observado; el verificador compara los campos persistidos y excluye ese indicador de sesión |
| Salida/reporte existente | Usar una ruta nueva, no `-Force`, borrados o sobrescrituras automáticas |
| Valores o snapshot no coinciden | Conservar evidencia, revisar calendario/modo/IDs. No ignorar la comparación para obtener `PASS` |
| Crear o validar consume minutos por tarea | Usar `Create -CreationMode Xml` y el snapshot XML predeterminado; no reconstruir un bucle de getters COM |
| Tabla más larga que la pantalla | `Export-ProjectImage.ps1`; no iniciar una serie de capturas por scroll |
| GIF/PNG no contiene la última fila | Revisar selección, memoria y rango; exportar bloques nativos si es necesario. `KeepRange` no garantiza memoria ilimitada |
| Duraciones cero al importar XML | Incluir `RemainingDuration` y `ActualDuration`, además de `Duration`, en formato de horas/minutos/segundos; el writer del kit ya lo hace |
| COM rechaza una ruta válida con argumento inválido | Convertir la ruta a `[string]` antes de pasarla a `InvokeMember`; un resultado enriquecido de PowerShell no siempre se convierte correctamente |
| Diario de vista en `Opening` | El intento pudo abrir un archivo antes de fallar. Revisar proyectos abiertos y diario; no relanzar `Open` automáticamente |
| Diario perteneciente a otro ProcessId | Project se reinició. No cerrar documentos basándose en una sesión anterior |
| `user input was detected` | Releer `get_window_state` antes de continuar; el usuario puede haber cambiado foco, selección o tamaño de ventana |
| La captura de Computer Use se detuvo por el usuario | Detener la interacción; no cambiar de capturador para continuar la acción cancelada |

## Escalamiento acotado

Para fallos puramente de lectura o selección de ventana, refresca el estado y permite un reintento si corresponde. Para mutaciones, no hay reintento automático. Conserva HRESULT, fase y rutas generadas sin volcar datos sensibles innecesarios.

La falta de una captura no autoriza reparar Office, cambiar drivers, elevar privilegios, desactivar antivirus o modificar seguridad. Si el bloqueo es del entorno/herramienta y no puede resolverse dentro del acceso autorizado, comunica el impedimento concreto.
