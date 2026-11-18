# Flujos y botones de PRPC (referencia para otras automatizaciones)

Este archivo resume los selectores y el orden de los pasos que usa `ui_automation.py`. Sirve para copiarlos en otra automatización de PRPC. Los valores vienen del script actual; si cambian en la app, hay que actualizarlos acá y en el script.

## Iframes

Pega carga la UI dentro de iframes cuyo nombre cambia (`PegaGadget1Ifr`, `PegaGadget2Ifr`...). Por eso el script no fija el nombre: busca el iframe que contiene el elemento que necesita.

- `PEGA_FRAME_PREFIX = "PegaGadget"`: prefijo de los iframes de Pega.
- `PEGA_FRAME_TIMEOUT_MS = 15000`: cuánto buscar el iframe.

Funciones que lo hacen (en `ui_automation.py`):
- `find_frame_with(page, selector)`: devuelve el `FrameLocator` del iframe que contiene el selector.
- `find_pega_frame(page)`: iframe que contiene el botón Actions.
- `find_frame_with(page, WORK_TAB)`: iframe que contiene la pestaña Work.

## Selectores

Todos los selectores de botones y opciones usan `>> visible=true`. Así, si hay elementos ocultos con el mismo texto, solo se clickea el visible.

| Constante | Selector | Uso |
|---|---|---|
| `SEARCH_INPUT` | `role=textbox[name='Search']` | Barra de búsqueda del JPV (está en la página principal, no en un iframe). |
| `WORK_TAB` | `text="Work" >> visible=true` | Pestaña Work del caso. Se clickea por texto exacto, como en codegen. |
| `WORK_PANEL` | `role=tabpanel[name='Work'] >> visible=true` | Contenido de la pestaña Work. Ahí figura el asignado. |
| `ACTIONS_MENU` | `role=button[name='Actions'] >> visible=true` | Botón Actions. |
| `CHANGE_CASE_TYPE_OPTION` | `role=menuitem[name='Change case type'] >> visible=true` | Opción del menú Actions para cambiar el tipo de caso. |
| `NEW_CASE_TYPE_LABEL` | `"Select New Case Type"` | Label del dropdown del nuevo tipo de caso. Se elige con `select_option(value=...)`. |
| `REASSIGN_OPTION` | `role=menuitem[name='Reassign'] >> visible=true` | Opción Reassign del menú Actions. |
| `STEP_BY_STEP_OPTIONS` | `role=menuitem[name=/^Step By Step/i]` y `role=button[name=/^Step By Step/i]` | Puede ser menuitem o botón, y puede tener un nombre de empleado después. |
| `PERFORM_RESEARCH_OPTION` | `role=menuitem[name=/^Perform research/i] >> visible=true` | Opción alternativa si no aparece Step By Step. Puede tener texto después. |
| `MESSAGE_ATTACHED_OPTIONS` | `role=menuitem[name=/^MessageAttached/i]` y `role=button[name=/^MessageAttached/i]` | Último recurso si no aparece ninguna opción anterior. |
| `REOPEN_BUTTON` | `role=button[name=/^Reopen/i] >> visible=true` | Si aparece, el caso no se reasigna (se saltea). |
| `OPERATOR_LABEL` | `"operator"` | Label del dropdown donde se elige el assignee (`select_option(label=...)`). |
| `SUBMIT_BUTTON` | `role=button[name='Submit'] >> visible=true` | Submit. |
| `CLOSE_BUTTON` | `role=button[name='Close'] >> visible=true` | Close. |

Notas:
- `MENU_CHECK_MS = 2000`: cuánto esperar a que aparezca una opción del menú Actions antes de decidir cuál usar.
- `POST_SEARCH_PAUSE_MS = 2000`: pausa después de buscar el JPV, antes de tocar el caso.

## Valores fijos del negocio

- `OWNERS_ALLOWED = ("FX12", "AutoCV")`: el reassign solo se hace si el campo Work contiene alguno de estos valores, como palabra completa.
- `BOT_ID = "B089707"`: si el campo Work contiene este ID, el caso está con el bot y se saltea.
- `CASE_TYPE_TRIGGER = "Miscellaneous"` y `NEW_CASE_TYPE = "JPM-Work-Shared-WDPC-Exp-AmendBen"`: si el tipo de caso es Miscellaneous y aparece Change case type, se cambia al nuevo tipo (se elige por `value`, no por texto visible).

## Orden de los pasos

1. **Buscar el JPV:** clic en `SEARCH_INPUT`, `fill(jpv)`, verificar que el valor quedó completo, Enter. Después esperar `POST_SEARCH_PAUSE_MS`.
2. **Pestaña Work:** clic en `WORK_TAB` (en el iframe que lo contiene). Esperar `STEP_PAUSE_MS`.
3. **Verificaciones (se saltean sin tocar el caso):**
   - Leer el texto de `WORK_PANEL`.
   - Si contiene `BOT_ID` como palabra completa: saltear (asignado al bot).
   - Si no contiene ninguno de `OWNERS_ALLOWED` como palabra completa: saltear.
   - Si aparece `REOPEN_BUTTON`: saltear.
4. **Actions:** clic en `ACTIONS_MENU`.
5. **Cambio de tipo (solo Miscellaneous y si aparece Change case type):**
   - Clic en `CHANGE_CASE_TYPE_OPTION`.
   - `select_option(value=NEW_CASE_TYPE)` en `NEW_CASE_TYPE_LABEL`.
   - Clic en Submit.
   - Volver a clickear `ACTIONS_MENU`.
6. **Reassign:** si `REASSIGN_OPTION` no es visible:
   - Si aparece `STEP_BY_STEP_OPTIONS` → clic en la que aparezca.
   - Si no, si aparece `PERFORM_RESEARCH_OPTION` → clic.
   - Si no, clic en la primera de `MESSAGE_ATTACHED_OPTIONS` que aparezca.
   - Esperar `STEP_PAUSE_MS`.
   - Clic en `REASSIGN_OPTION`.
7. **Elegir el assignee:** `select_option(label=assignee)` en `OPERATOR_LABEL`.
8. **Cerrar:** clic en Submit, luego en Close.

## Funciones reutilizables

En `ui_automation.py`:
- `contiene_valor(texto, valor)`: busca el valor como palabra completa (`FX1213` no es `FX12`).
- `menu_option_visible(pega, selector)`: `True` si el elemento aparece en `MENU_CHECK_MS`.
- `find_frame_with(page, selector)`: iframe de Pega que contiene el selector.
- `SkipRow`: excepción para saltear una fila sin contarla como error.
- `guardar_resultados(results)`: escribe `resultados.csv` y muestra la tabla de errores y el resumen.

## Login

- `login_sso(page)`: clic en **Login with SSO** y en **JPMorgan Chase (Default)**. Usuario, contraseña y MFA los completa la persona.
- Después el script espera hasta `LOGIN_TIMEOUT_MS` a que aparezca `SEARCH_INPUT`.
