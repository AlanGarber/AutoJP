# Conocimiento del proyecto AutoJP

Documento de contexto para retomar el trabajo en este repo desde cualquier sesión de Claude (incluida la de la VDI), sin depender del historial de chats. Resume las decisiones, el estado de cada automatización y los pendientes.

Fecha de corte: 2026-11-18. Estado del código: `main` en el commit de la búsqueda por pasos chicos (ver historial). Si el código cambió después, manda el código: este documento puede quedar desactualizado.

Datos sensibles: no agregar acá nombres de compañeros, correos, IDs de casos ni rutas personales. Los valores de negocio están en la sección `CONFIGURACIÓN` de cada script.

---

## 1. Objetivo

Automatizar tareas repetitivas del flujo diario del equipo, con UI web (Playwright) o Outlook de escritorio (PowerShell). Cada automatización vive en su propia carpeta con su propio README. El índice está en el `README.md` de la raíz.

## 2. Estado de cada automatización

| Carpeta | Qué hace | Estado |
|---|---|---|
| `asignar-casos-empleados` | Asigna casos a empleados en Pega desde un XLSX, fila por fila. | Código completo, con selectores ajustados con codegen y DevTools. Hay un problema abierto con el botón Actions (ver sección 9). |
| `reasignar-casos-pega-lob` | Reasigna casos de Pega a la work queue de cada LOB, desde un XLSX. | Borrador, en uso. `DRY_RUN = False` desde 2026-11-18 (reasigna de verdad). Caso especial "To check" en pausa desde 2026-11-18. Varios valores `PLACEHOLDER` sin confirmar. |
| `mails-sin-responder` | Reporte diario de mails de casillas compartidas de Outlook que nadie del equipo respondió. | En uso. Falta completar `$Equipo`. Solo lectura. |

## 3. asignar-casos-empleados

### Qué hace
Por cada fila del XLSX, busca el JPV en la search bar y reasigna el caso al assignee indicado.

### Archivo de entrada (XLSX)
- Columna A: assignee (persona a la que se asigna).
- Columna B: JPV (identificador del caso, lo que se busca).
- Columna C: equipo. Solo se procesan los equipos de `TEAMS_ALLOWED`; el resto queda `OMITIDA`.
- Columna D: tipo de caso. Los tipos `Miscellaneous` y `OFAC` solo se procesan si el equipo es AUTOCV (regla del 2026-11-19); para otros equipos la fila queda `OMITIDA` sin tocar el caso. Excepción desde 2026-11-20: `iFX` + `Miscellaneous` se procesa con Change case type a Amendment. Sin distinguir mayúsculas.
- Las demás columnas se ignoran.
- Hoja: `Cases Distribution` o `Distribution` (`XLSX_SHEETS`, nombres exactos, en ese orden). Desde 2026-11-20 se acepta también `Distribution`, porque llegó un XLSX con esa hoja. Si no está ninguna, el script se detiene antes de abrir el navegador (regla del 2026-11-19). Con encabezado (`XLSX_HAS_HEADER`).

### Archivo failed (2026-11-20)
- Se reconoce porque el nombre del archivo contiene `failed` (`FAILED_FILE_MARK`).
- Columna A: Investigator (assignee). Columna B: JPV. Columna E: Sub LOB. No tiene equipo ni tipo de caso.
- La columna E solo tiene dos valores: `12` y `13`. Solo se procesan las filas con `12` (`FAILED_SUB_LOB`), confirmado por Alan el 2026-11-20. "Por ahora": si cambia, se cambia esa constante.
- Hoja: `Distribution` (confirmado por Alan el 2026-11-20).
- Las demás reglas (Work, owner, Reopen, Reassign) son las mismas que para el archivo normal. Pendiente de confirmar con Alan si aplican igual.

### Flujo por fila
1. Escribe el JPV en la search bar y confirma con Enter (`SEARCH_WITH_ENTER`). Espera `POST_SEARCH_PAUSE_MS`.
2. Abre la pestaña **Work** y verifica el caso. Si la fila no cumple, se saltea sin tocar el caso:
   - Si el asignado actual contiene el ID del bot (`BOT_ID`), se saltea.
   - Si el campo Work no contiene ninguno de los owners permitidos (`OWNERS_ALLOWED`), se saltea.
   - Si el campo Work dice `AutoCV Automation`, se saltea aunque también diga `AutoCV` (`OWNER_NO_REASIGNAR`, regla del 2026-11-20). Se reasigna solo si dice `AutoCV`.
   - Si aparece el botón de reabrir (`REOPEN_BUTTON`), se saltea.
2a. Si Actions tiene **ReviewPymtPaidCorrectly**, el flujo es: Actions > ReviewPymtPaidCorrectly > Actions > **Cancel this Assignment** > Submit > Actions, y sigue el flujo normal. Pedido por Alan el 2026-11-24.
2b. Optimización del 2026-11-24: una sola espera para que aparezca el menú Actions (en vez de 2 s por cada opción que no está). Pausas bajadas: `POST_SEARCH_PAUSE_MS` 1000 y `STEP_PAUSE_MS` 200. El log muestra el tiempo por fila y los pasos más lentos. Si vuelve a tocar el caso anterior o falla un click, subir las pausas. Reopen sigue esperando 2 s a propósito: un falso negativo reasignaría un caso reabierto.
3. Si el tipo de caso es `CASE_TYPE_TRIGGER` y aparece **Change case type** en Actions, cambia el tipo a `NEW_CASE_TYPE` (se elige por *value*, no por texto visible), hace Submit y vuelve a abrir Actions.
4. **Reassign**: si la opción no está en el menú, prueba en este orden **MessageAttached**, **Step By Step** y **Perform research**. Luego clickea **Reassign**. Cambio pedido por Alan el 2026-11-23: MessageAttached va primero, porque es el flujo directo.
5. Elige el assignee en el dropdown **operator**. Si el nombre del XLSX no coincide con el de la app, se usa `NAME_ALIASES`.
6. **Submit** y **Close**. No hace falta Save.
7. Espera `PAUSE_BETWEEN_ROWS_MS` antes de la siguiente fila. Se decidió una pausa mínima porque el script iba demasiado rápido.

### Login
- `AUTO_SSO_CLICKS = True`: el script hace los clicks de SSO (**Login with SSO** y **JPMorgan Chase (Default)**).
- Usuario, contraseña y MFA los completa la persona. El script espera hasta `LOGIN_TIMEOUT_MS` a que aparezca la búsqueda.

### Errores y recuperación
- Si una fila falla, guarda `error_fila_N.png`, presiona Escape, recarga y vuelve a la búsqueda.
- Si no logra volver a la búsqueda, detiene la corrida y marca las filas restantes como `NO PROCESADA`.
- Un assignee que no existe en el dropdown deja la fila en `ERROR` y el script sigue.

### Salida: `resultados.csv`
Estados: `OK`, `ERROR`, `OMITIDA` (equipo fuera de `TEAMS_ALLOWED`) y `NO PROCESADA` (la corrida se detuvo antes).

### Uso (VDI)
Se corre desde la carpeta `asignar-casos-empleados` con el XLSX del día, por ejemplo `python ui_automation.py "Distribution.xlsx"`. El archivo de la mañana y el de la tarde tienen nombres distintos; el de *failed* se reconoce porque su nombre contiene `failed` (`FAILED_FILE_MARK`) y se lee con sus propias columnas (ver sección 3, "Archivo failed"). Las instrucciones paso a paso para personas no técnicas están en `INSTRUCCIONES.MD`.

### Documentación de apoyo
- `FLUJOS_PEGA.md` (en la carpeta): selectores, orden de pasos y funciones reutilizables de PRPC. Sirve de referencia para otras automatizaciones de Pega.

## 4. reasignar-casos-pega-lob (borrador)

### Qué hace
Reasigna cada caso a la work queue que corresponde a su LOB.

### Archivo de entrada (XLSX)
- Columna A (**Comments**): LOB.
- Columna B: se usa solo en el caso especial.
- Columna C (**ID**): JPV, que es el ID del caso.

### Reglas de negocio
- `Corp&FI` → work queue `WDPCTSTWRKFX12`.
- `Retail` → work queue `WDPCTSRWRKFX13`.
- Work queue FX13 (2026-11-18): el value correcto es `WDPCTSRWRKFX13` (con `SR`, no `ST`). Lo confirmó Alan en la app. FX12 sigue siendo `WDPCTSTWRKFX12`.
- **Caso especial (en pausa desde 2026-11-18):** si A es `To check` **y** B es `FX Retail Inv` a la vez, iría a `WDPCTSRWRKFX13`. Mientras esté en pausa (`CASO_ESPECIAL_ACTIVO = False`), esas filas se omiten. Si solo se cumple una de las dos, la fila se omite.
- El caso solo se reasigna si su case owner es exactamente `FX1213`. Hoy la validación busca ese valor como palabra completa dentro del panel Work; falta pasarla a igualdad exacta con el locator del campo **Case owner**.

### Flujo
1. La basket se abre como popup desde el dropdown **View Queue for** de la ventana principal (confirmado con codegen el 2026-11-18). Hace click en el JPV dentro de la basket.
2. El caso se abre en la ventana principal de Pega, dentro de iframes `PegaGadget*`.
3. Pestaña Work, verificación del owner.
4. **Actions** > **Reassign** (la opción está dentro del menú Actions, es un menuitem). Si Reassign no aparece en el menú, **MessageAttached** y después **Reassign**; si tampoco, **Perform research** (con el nombre del empleado después) y **Reassign**. Orden actualizado el 2026-11-23 (antes era Perform research primero, desde 2026-11-18).
5. Elige la work queue en el dropdown **Work Queue** (`label=Work Queue`, por `value`).
6. **Submit** y **Close**. En modo simulación, **Close** y **Discard** (no se guarda nada).

### Salida: `resultados.csv`
Estados: `SIMULADO` (DRY_RUN), `OK`, `ERROR`, `OMITIDA` (owner distinto de `FX1213`, sin work queue según A y B, o JPV vacío) y `NO PROCESADA`.

### Configuración
- `DRY_RUN = False` desde 2026-11-18: el script hace Submit y Close de verdad. Antes estaba en `True` (Close y Discard). Se cambió a pedido del usuario, con los selectores PLACEHOLDER todavía sin validar.
- `ZOOM_BASKET = 0.5` (2026-11-18): la basket se achica al 50% por CSS al cargarla, para que entren más filas y el scroll sea más corto. Alan confirmó que el zoom manual funciona. Si se hace Ctrl + - a mano en el mismo sitio, el efecto se suma: poner `None`.
- `JPV_LINK_NTH = 2`: índice del link del JPV en la basket, según codegen. Confirmado el 2026-11-18.
- `CASO_ESPECIAL_ACTIVO = False`: pausa del caso "To check". Decidido el 2026-11-18 para testear.
- Ventana maximizada (2026-11-18): con el navegador en tamaño chico no aparece el dropdown **View Queue for**. Por eso el script abre el navegador maximizado (`MAXIMIZAR_VENTANA = True`).
- Grilla de la basket (2026-11-18): tiene muchos JPV y hay que bajar para verlos. El scroll se hace por JavaScript sobre los contenedores con scroll, no con la rueda del mouse (la rueda no llegaba a la grilla). Desde 2026-11-18 recorre la grilla desde arriba en pasos chicos y revisa después de cada paso hasta el final; se probó sobre una grilla de prueba local. El click va sobre el link visible del JPV (no sobre el elemento número 2, que puede no ser el visible). El script busca bajando y, si no aparece, desde el inicio. La basket se maximiza por CDP (`Browser.setWindowBounds`). Pendiente de verificar contra Pega.
- Instrucciones (2026-11-19): `INSTRUCCIONES.MD` cubre las dos automatizaciones, cada una en su sección. La B (reasignar-casos-pega-lob) tiene instrucciones de uso diario, separadas de la A.
- Excel (2026-11-19): el XLSX se lee directo de la carpeta de red del equipo, sin copiarlo a la carpeta del script. La ruta va en `XLSX_RUTA` de `config_local.py` y no se sube al repo. Sin argumento, el script usa esa ruta; con argumento, usa el archivo pasado.
- Actions (2026-11-23): en la automatización B, si Reassign no está en el menú, primero se prueba **MessageAttached** (Reassign dentro), y después **Perform research**. Mismo orden que en la A, pedido por Alan. Finalizar sigue igual: Submit y Close.
- Modo prueba (2026-11-19): `--buscar JPV` busca un JPV en la basket sin hacer clic ni reasignar. Se usa para probar antes de tener el Excel.
- Basket (2026-11-18): la grilla tiene paginación al final (botones Next y Previous), confirmado por Alan. Cada búsqueda de JPV empieza en la página 1, recorre la página de arriba hacia abajo y pasa a la siguiente hasta la última. Si no aparece, la fila se omite y la siguiente fila arranca otra vez desde la página 1. Puede estar dentro de un iframe de la basket, así que el script busca y hace scroll en la página y en todos sus frames. Los selectores de Next y Previous son suposiciones, a validar en la VDI. Si un JPV no aparece, `diagnostico_basket.txt` indica en qué frame está el texto.
- Case owner (2026-11-18): el valor que tiene que figurar en el caso para reasignarlo pasó a ser `FX1213`. Antes se había definido como `TSFX1213`.
- Criterio de omisión (2026-11-18): si el JPV no aparece en la basket, la fila queda `OMITIDA` y el script sigue con la siguiente.
- Login y basket automáticos (2026-11-18): el script hace los clicks de SSO, espera el Enter después del login y abre la basket (My teams > View Queue for). `BASKET_QUEUE_VALUE = "WDPCTSXWRKFX1213"`, confirmado el 2026-11-18.
- `WORK_QUEUE_SELECTOR = 'internal:label="Work Queue"i'` (mismo criterio que `get_by_label` del codegen).

## 5. mails-sin-responder

### Qué hace
Para un día (por defecto hoy, o el que se pase con `-Fecha`), lee los mails recibidos en casillas compartidas de Outlook y arma un reporte con los que nadie del equipo respondió. **Solo lee**: no marca, mueve ni borra mails.

### Requisitos
- Outlook de escritorio (clásico) abierto en la VDI, con permiso de lectura sobre las casillas. El nuevo Outlook no permite que otros programas lean los mails.
- PowerShell. Si la VDI bloquea scripts `.ps1`, el README tiene la alternativa.

### Cómo se decide que un mail está respondido (mismo día)
1. Outlook lo marca como respondido o respondido a todos (vale sin importar quién respondió).
2. Hay un mail posterior en la misma conversación, en los Enviados de la casilla o en los de la persona que corre el script.
3. Llegó a la casilla un mail de alguien de `$Equipo` en la misma conversación, después del original (respuesta a todos con la casilla en copia).

Se ignoran: mails del propio equipo, respuestas automáticas, avisos de entrega, invitaciones, remitentes tipo `noreply`, y los remitentes y asuntos de `$RemitentesIgnorados` y `$AsuntosIgnorados`.

### Configuración (sección `CONFIGURACIÓN` del `.ps1`)
- `$Casillas` (obligatorio): casillas compartidas tal como aparecen en Outlook. Se puede agregar `'PERSONAL'`.
- `$Equipo` (recomendado): correos SMTP de quienes responden desde las casillas. **Hoy está vacío.**
- `$IncluirSubcarpetas = $true`: revisa también las subcarpetas de la Bandeja de entrada.
- `$SoloSiEstaEnPara = $false`, `$ContarReenvioComoRespuesta = $false`, `$AgruparPorConversacion = $true`.

### Salida
- `reporte_no_respondidos_AAAA-MM-DD.html`: resumen por casilla y detalle. Se abre al terminar.
- `reporte_no_respondidos_AAAA-MM-DD.csv`: detalle para Excel (casilla, hora de recepción, remitente, asunto, carpeta, si estaba en Para o CC, horas sin respuesta y mails del hilo sin responder).

### Uso
Abrir Outlook y hacer doble click en `ejecutar.bat`. Para otro día: `ejecutar.bat -Fecha AAAA-MM-DD`.

### Límites conocidos
- No detecta una respuesta de un compañero desde su casilla personal, sin la compartida en copia y sin responder el mail original.
- No detecta una respuesta enviada como mail nuevo con otro asunto.

### Diagnóstico de la primera corrida
Al revisar un caso concreto se vio que casi todas las respuestas del equipo llegan como mails recibidos en la casilla compartida. Con `$Equipo` vacío, esas respuestas no cuentan y los mails quedan como pendientes. Por eso el reporte marcaba como pendientes casi todos los casos que sí tenían respuesta. La corrección es completar `$Equipo`.

## 6. Reglas de trabajo acordadas

- **Commits y push:** solo cuando la persona lo pide explícitamente. Mensajes de commit en español, cortos y descriptivos.
- **Cambios terminados:** se mergean siempre a `main` (fast-forward) y se pushean. Esta regla vale para todo cambio futuro en AutoJP.
- **Ejecución contra la app real:** no correrla sin confirmación. El usuario la ejecuta en su propia terminal (VDI), porque pide login SSO y MFA.
- **Verificación mínima:** `python -m py_compile <script>` antes de dar un cambio por listo. Si Playwright está instalado, probar con un XLSX de prueba local.
- **Stack:** Python 3.9+, Playwright (API sincrónica), `openpyxl` para leer XLSX. Sin dependencias nuevas sin preguntar.
- **Entrada:** siempre `.xlsx`. No se aceptan CSV ni XLS.
- **Datos sensibles:** los XLSX (`*.xlsx`), `resultados.csv`, `error_fila_*.png`, los reportes de Outlook y `config_local.py` están en `.gitignore` y no se suben. No pegar datos reales en commits, issues ni servicios externos.
- **Credenciales:** nunca guardar usuarios, contraseñas, tokens ni cookies en el repo.
- **Instrucciones para el equipo:** que bajen el repo con `git clone`, no con ZIP.
- **Idioma:** todo el contenido en español rioplatense (voseo). Los nombres de variables y funciones de Python, en inglés.
- **Preguntas al usuario:** una sola pregunta concreta cuando falta un dato. Confirmar antes de algo difícil de revertir (push, borrar, cambiar el formato de entrada).

## 7. Instalación y configuración (resumen)

1. `git clone https://github.com/AlanGarber/AutoJP.git`.
2. `pip install playwright openpyxl`.
3. Usar el navegador ya instalado: `BROWSER_CHANNEL = "msedge"` (o `"chrome"`). Si no hay navegador, `playwright install chromium` (puede fallar por la red corporativa; en ese caso, pedir el proxy a IT).
4. Copiar `config_local.example.py` como `config_local.py` y completar `APP_URL`. Este archivo no se sube.
5. Sacar selectores con codegen: `playwright codegen --channel msedge <URL>`. Si la app es interna, hace falta la VDI o la VPN.

## 8. Historial de decisiones (resumido)

- **Tecnología:** se eligió Python con Playwright en lugar de Selenium, por la espera automática y por `codegen`. Si la app fuera de escritorio o Citrix, Playwright no serviría y habría que usar pyautogui o pywinauto.
- **Repo:** se creó AutoJP y se reorganizó en una carpeta por automatización, con un README raíz como índice.
- **Renombrado:** la primera automatización se llamaba `ui-desde-csv` y pasó a `asignar-casos-empleados`.
- **Entrada:** el archivo cambió de CSV a XLSX. Primera columna, assignee; segunda, JPV.
- **Login:** se agregaron los clicks automáticos de SSO. Usuario, contraseña y MFA los completa la persona.
- **Ritmo:** se agregó una pausa entre filas porque el script iba demasiado rápido.
- **Pega:** se diagnosticó la estructura. La basket es un popup y el caso se abre en la ventana principal, con los controles dentro de iframes `PegaGadget*`. Por eso el script trabaja con dos páginas.
- **2026-11-18, reasignar-casos-pega-lob:** se confirmó con codegen el flujo de work queue. Retail → `WDPCTSRWRKFX13` y el índice del JPV en la basket quedaron confirmados. Se decidió que el caso especial "To check" queda en pausa para testear el resto del flujo.
- **2026-11-19, asignar-casos-empleados:** los tipos `Miscellaneous` y `OFAC` solo se reasignan si el equipo es AUTOCV. Para otros equipos, la fila se omite sin tocar el caso. Pedido de Alan.
- **Outlook:** `mails-sin-responder` se creó y se revisó su código. Quedó como solo lectura.
- **Instrucciones para el equipo:** se quitó la opción de ZIP. Se usa solo `git clone`.

## 9. Pendientes abiertos

### asignar-casos-empleados
- Problema con el botón **Actions** en el segundo caso: su estructura no coincide con el selector. Quedó pendiente de revisar el HTML o el codegen de ese botón.
- Archivo de *failed* implementado (2026-11-20): hoja `Distribution`, Sub LOB `12` (los valores son solo `12` y `13`). Falta probarlo en la VDI.

### reasignar-casos-pega-lob
- Cambiar la validación del owner a igualdad exacta, con el locator del campo **Case owner** en la pestaña Work.
- Confirmar los selectores `PLACEHOLDER` contra la app real.
- Caso especial "To check" en pausa (`CASO_ESPECIAL_ACTIVO = False`). Antes de reactivarlo, confirmar que TSFX13 = `WDPCTSRWRKFX13` y si la validación de owner aplica también a ese caso.
- Validar selectores y el campo Case owner en la VDI, con casos de prueba controlados (DRY_RUN ya está en `False`).

### mails-sin-responder
- Completar `$Equipo` con los correos de quienes responden desde las casillas. Para copiarlos, abrir un mail de cada persona y tomar la dirección del campo "De" (no el nombre).
- Volver a correr el mismo día y comparar cuántos pendientes bajan.
- Propuesta sin implementar: agregar contadores de diagnóstico al reporte (cuántos Enviados se leyeron por casilla, cuántos mails tienen marca de respondido, cuántos se resolvieron por marca y cuántos por hilo). Sirve para ver dónde se pierden las respuestas.

## 10. Cómo retomar el trabajo

1. Leer este documento, el `README.md` raíz y `CLAUDE.md`.
2. Leer el README de la automatización que se va a tocar.
3. Antes de tocar código, verificar el estado con `git status` y `git log`.
4. Al terminar un cambio: `py_compile`, commit en español, y merge a `main` solo si la persona lo pidió.
