# Reporte de mails sin responder

Revisa una o varias casillas compartidas (y, si querés, tu casilla personal), lee los mails que entraron en el día y arma un reporte con los que nadie del equipo respondió.

Un mail cuenta como **respondido** si, ese mismo día:

1. Outlook lo marca como respondido (Responder o Responder a todos). La marca queda en el mail de la casilla, así que vale sin importar **quién** del equipo respondió.
2. Hay un mail en la misma conversación, enviado después, en los Enviados de la casilla o en los tuyos.
3. Llegó a la casilla un mail de alguien del equipo en la misma conversación, después del original. Pasa cuando un compañero responde a todos con la casilla en copia.

Se ignoran los mails del propio equipo, las respuestas automáticas, los avisos de entrega, las invitaciones a reuniones y los remitentes tipo `noreply`.

Todo está en un solo archivo: `reporte_no_respondidos.ps1`.

## 1. Instalación

No hay que instalar nada. Usa PowerShell, que viene con Windows, y el Outlook de escritorio (clásico) que ya está en la VDI. No hace falta pedir software, acceso a APIs ni abrir Outlook en el navegador.

Lo único que necesitás es **permiso de lectura sobre las casillas compartidas**, el mismo que ya usás para verlas en Outlook.

Copiá la carpeta `mails-sin-responder` a la VDI (por ejemplo, al Escritorio). Si no podés pasar archivos, creá los archivos a mano en el Bloc de notas y pegá el contenido. Guardá el `.ps1` con codificación **UTF-8 con BOM** (en el Bloc de notas: **Guardar como > Codificación: UTF-8 con BOM**) para que no se rompan las tildes.

## 2. Configurar el script

Abrí `reporte_no_respondidos.ps1` con el Bloc de notas y completá la sección `CONFIGURACIÓN`:

- `$Casillas` (**obligatorio**): el email o el nombre de cada casilla compartida, tal como aparece en Outlook. Agregá `'PERSONAL'` si querés incluir tu casilla.

  ```powershell
  $Casillas = @(
      'operaciones@empresa.com'
      'Cobranzas Equipo'
  )
  ```

- `$Equipo` (recomendado): los emails de las personas que responden desde las casillas. Así, si alguien contesta con la casilla en copia, cuenta como respondido, y los mails internos entre ustedes no aparecen como pendientes.
- `$IncluirSubcarpetas`: revisar también las subcarpetas de la Bandeja de entrada y de Elementos enviados.
- `$SoloSiEstaEnPara`: `$true` para ignorar los mails donde la casilla está solo en CC o llegó por una lista.
- `$ContarReenvioComoRespuesta`: `$true` para que reenviar un mail cuente como respuesta.
- `$AgruparPorConversacion`: `$true` para ver una sola fila por hilo.
- `$RemitentesIgnorados` y `$AsuntosIgnorados`: sumá remitentes o asuntos que nunca esperan respuesta (alertas, newsletters, etc.).

## 3. Ejecutar

1. Abrí Outlook en la VDI.
2. Hacé doble click en `ejecutar.bat`.

Para un día anterior, desde una terminal en esta carpeta:

```bat
ejecutar.bat -Fecha 2026-11-12
```

Al terminar se abre el reporte en el navegador y quedan dos archivos en la carpeta:

- `reporte_no_respondidos_AAAA-MM-DD.html`: resumen por casilla y detalle, para leer. Tiene botones para filtrar por *Para nosotros*, *En copia* o *Lista/otro*, por casilla y por carpeta.
- `reporte_no_respondidos_AAAA-MM-DD.csv`: el detalle, para abrir en Excel.

Columnas del detalle: casilla, hora de recepción, remitente, email, asunto, carpeta, si la casilla estaba en Para/CC, horas sin respuesta y cuántos mails del hilo quedaron sin responder.

## Asignado

Si el asunto tiene un JPV (`JPV` + dígitos + `-` + dígitos), el script busca el assignee en la base `$ArchivoCasos` (columna A = JPV, columna Q = assignee). La ruta real va en `asignacion_local.ps1`, fuera del repo; la plantilla es `asignacion_local.example.ps1`.

- Si el assignee es una cola (`$AssigneesDeCola`: FX12, AutoCV, AutoCV Automation), la columna muestra el responsable de la casilla en ese día y horario, según `$ArchivoHorarios` (filas 4 a 7, columnas D a H = lunes a viernes). Si no hay responsable en ese horario (fuera de horario, fin de semana, celda vacía o horario no cargado), muestra la cola tal como está en el Excel.
- Si el JPV no está en la base, la columna dice `JPV no encontrado`.
- Si el asunto no tiene JPV, la columna dice `Sin JPV en el asunto`. Si no hay base configurada, queda vacía.

## Si algo falla

- **"No se puede cargar el archivo... la ejecución de scripts está deshabilitada"**: la VDI bloquea scripts `.ps1`. Abrí PowerShell en esta carpeta y corré:

  ```powershell
  & ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 .\reporte_no_respondidos.ps1)))
  ```

- **"Se saltea la casilla..."**: el nombre o el email en `$Casillas` no coincide con ninguna casilla, o no tenés permiso. Copiá el email exacto desde Outlook.
- **"No pude abrir los Enviados de..."**: tenés permiso sobre la Bandeja de entrada de la casilla pero no sobre Enviados. El reporte igual funciona con la marca de "Respondido" de cada mail.
- **"No pude conectarme a Outlook"**: abrí Outlook antes de correr el script y verificá que sea el Outlook clásico (el "nuevo Outlook" no permite que otros programas lean los mails). Si PowerShell está en modo restringido, este comando muestra `ConstrainedLanguage`:

  ```powershell
  $ExecutionContext.SessionState.LanguageMode
  ```

  En ese caso, PowerShell no puede usar Outlook y hay que hablar con IT.
- **Outlook muestra "Un programa está intentando acceder a la información de correo"**: elegí **Permitir**. Si no aparece la opción, la política de la empresa lo bloquea.

## Notas

- El script solo lee: no mueve, marca ni borra mails.
- Los reportes tienen remitentes y asuntos reales. Están en `.gitignore` y no se suben al repo.
- Si la casilla no aparece en el panel de carpetas de Outlook, el script abre solo la Bandeja de entrada principal, sin las subcarpetas. Para incluirlas, agregá la casilla a tu Outlook (**Archivo > Configuración de la cuenta > Más configuraciones > Avanzadas > Agregar buzón**).
- No se detecta una respuesta de un compañero que salga desde su casilla personal sin la casilla compartida en copia y sin responder desde el mail original. Tampoco una respuesta enviada con otro asunto como mail nuevo.
