# Capturas reales de Project

## Qué está probado

COM creó/modificó archivos y una captura nativa de Windows guardó la ventana visible en PNG. En la sesión original, Computer Use enumeraba ventanas, las activaba y leía controles, pero no capturaba frames. El clic por índice también falló por falta de geometría. La captura GDI resolvió la evidencia visual sin reparar ni reinstalar Project.

Este helper no es una reparación del plugin. Solo captura la ventana **en primer plano**, mediante `System.Drawing.Graphics.CopyFromScreen`. No sirve para trabajar con Windows bloqueado, ni evita restricciones de seguridad. Si una herramienta rechaza la captura por permisos, o el usuario la detiene, no uses este mecanismo para eludirlo.

## Secuencia completa

Usa rutas absolutas. En los comandos, `$scripts` es la carpeta de scripts de esta skill y `$runDir` la carpeta de la ejecución, como se definen en el README.

### 1. Guardar y verificar los cambios

Termina `Create` o `UpdateCopy` y exige su reporte `PASS`. No captures mientras otra operación COM sigue ejecutándose. Para un «antes/después», captura primero el archivo original en lectura, restaura la vista, crea la copia editada y repite sobre ella. Usa nombres distintos (`before.png`, `after.png`).

### 2. Abrir el resultado para verlo

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-View.ps1" `
  -Action Open -ProjectPath "$runDir\edited.mpp" -SessionPath "$runDir\view.json"
```

El helper conserva la identidad del documento anterior. Abre el resultado en lectura si no estaba abierto y registra si es dueño de esa apertura. Si ya estaba abierto, solo lo activa y no lo cerrará al restaurar. El diario en JSON contiene rutas locales: déjalo en `runs/`.

### 3. Activar la ventana correcta

Si Computer Use está disponible, lee una vez su skill y sus instrucciones de Windows de la versión instalada. Utiliza su entrada `node_repl` + `@oai/sky`; no deduzcas que está roto porque otro runtime (`cua_repl`) no exponga APIs nativas. Los nombres concretos de herramientas los determina el entorno.

Primera celda, sin elegir todavía coordenadas o índices:

```javascript
if (!globalThis.sky) {
  const { sky } = await import('@oai/sky');
  globalThis.sky = sky;
}
globalThis.projectWindows = (await sky.list_windows())
  .filter(w => /winproj/i.test(w.app));
nodeRepl.write(JSON.stringify(projectWindows));
```

Observa el resultado y selecciona la ventana principal del archivo solicitado. Puede haber una ventana principal y ventanas hijas para el mismo documento. No elijas el primer resultado sin comprobarlo; no copies un ID de esta guía o de otra sesión. En otra celda, usa exactamente un objeto `Window` devuelto:

```javascript
// Sustituye el titulo por el titulo principal observado en la celda anterior.
globalThis.projectMatches = projectWindows.filter(
  w => w.title === 'TITULO PRINCIPAL OBSERVADO'
);
if (projectMatches.length !== 1) throw new Error('Seleccion de ventana ambigua');
globalThis.projectWindow = projectMatches[0];
await sky.activate_window({window: projectWindow});
```

Si quieres confirmar controles, pide `get_window_state({window: projectWindow, include_screenshot:false, include_text:true})`; no añadas una captura del plugin solo por rutina si ya sabes que falla. Los handles/índices deben refrescarse si cambia el estado.

Sin Computer Use, el usuario puede poner Project delante. La captura se ejecuta desde una herramienta de shell en segundo plano; abrir una consola visible para lanzarla puede quitarle el foco a Project. No uses una automatización alternativa prohibida por las reglas de tu entorno para forzar la activación.

### 4. Capturar

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Guardar-Captura-Project.ps1" `
  -OutputPath "$runDir\after.png" -ExpectedTitle 'edited.mpp'
```

Se puede añadir `-WindowHandle` con el ID nativo recién observado. El script comprueba proceso `WINPROJ`, primer plano, título esperado, ventana no minimizada y rectángulo visible. Rechaza sobrescribir un PNG existente. Si cambia el foco durante la lectura, no guarda la imagen.

El título identifica el nombre del documento, no necesariamente su ruta completa. Si hay archivos homónimos, confirma también `ActiveProject.FullName` mediante COM antes de activar/capturar; el paso `Project-View Open` ya comprueba esa ruta al abrir.

### 5. Abrir el PNG y comprobarlo

Usa la herramienta de imágenes locales del entorno, por ejemplo `view_image`. Revisa:

- ¿Corresponde al archivo y a la etapa correctos?
- ¿Se leen las tareas y columnas solicitadas? ¿El Gantt está dentro del área visible?
- ¿Aparecen el cambio esperado y las fechas? Una columna que no se ve no constituye evidencia de ese dato.
- ¿Hay un diálogo, otra ventana, notificación o contenido personal que impida usar la imagen?

Una captura puede tener poca anchura si la aplicación está ajustada a media pantalla. Ampliar la ventana no siempre cambia su tamaño efectivo y puede cambiar el foco: vuelve a observar y activar. Si falta una columna, prepara la vista sobre la copia o entrega la tabla verificada junto con la imagen; no afirmes que el PNG muestra datos que no aparecen.

No generes una imagen artificial ni alteres cifras para simular evidencia. Si recortas por un requisito explícito, identifícala como recorte y conserva la captura original localmente. No incluyas capturas reales con cuentas visibles en el repositorio de instrucciones.

### 6. Restaurar

```powershell
powershell.exe -NoProfile -NonInteractive -STA -File "$scripts\Project-View.ps1" `
  -Action Restore -SessionPath "$runDir\view.json"
```

Solo se cierra la vista abierta por el helper, sin guardar cambios de presentación. Se comprueba que el archivo conserve su hash y que la instancia de Project sea la misma. Si el diario está en `Opening`, la aplicación se reinició o falta el documento original, inspecciona el estado; no fabriques un diario nuevo ni cierres todos los documentos para continuar.

Referencias: [Computer Use en Windows](https://learn.chatgpt.com/docs/computer-use) y [Graphics.CopyFromScreen](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.graphics.copyfromscreen).
