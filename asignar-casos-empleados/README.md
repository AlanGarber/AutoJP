# Asignar casos a cada empleado según XLSX

Reasigna casos a empleados en Pega a partir de un archivo XLSX. Por cada fila:

1. Busca el JPV (columna B) en la search bar y confirma la búsqueda.
2. Si el tipo de caso (columna D) es **Miscellaneous** y aparece **Change case type** en el menú **Actions**: elige **Amendment** en el dropdown **Select New Case Type** y hace click en **Submit**. Si no aparece, sigue con Reassign.
3. Abre **Actions** y elige **Reassign**. Si **Reassign** no aparece en el menú, elige **Step By Step:**; si tampoco aparece, elige **Perform research:**. Después elige **Reassign**.
4. Elige el assignee (columna A) en el dropdown **operator**.
5. Hace click en **Submit** y **Close**. No hace falta **Save**: al reasignar no hay cambios que guardar.

Al inicio, el script hace los clicks de SSO (**Login with SSO** y **JPMorgan Chase (Default)**). Usuario, contraseña y MFA los completás vos. Solo se procesan las filas de los equipos permitidos (columna C); el resto de las columnas se ignoran. Todo está en un solo archivo: `ui_automation.py`.

## 1. Instalación

Necesitás Python 3.9 o superior.

```bash
pip install playwright openpyxl
playwright install chromium
```

Si `playwright install chromium` falla con `ETIMEDOUT` o un error de descarga, la red corporativa está bloqueando la descarga. Dos opciones:

- **Usar Chrome o Edge ya instalados (recomendado):** en `ui_automation.py` poné `BROWSER_CHANNEL = "msedge"` (o `"chrome"`). No hace falta descargar nada. Para codegen, usá `--channel chrome` o `--channel msedge` (ver paso 2).
- **Configurar el proxy y reintentar:** en PowerShell, `$env:HTTPS_PROXY="http://PROXY:PUERTO"` y después `playwright install chromium`. La dirección del proxy la pide IT.

## 2. Sacar los selectores con codegen

Playwright tiene un grabador que escribe el código por vos.

```bash
playwright codegen --channel msedge https://TU-APP-INTERNA
```

Si no usás Edge, cambiá `msedge` por `chrome`. Si no tenés ninguno de los dos, quitá `--channel` y usá el Chromium de Playwright, que requiere el paso de descarga de la sección 1.

Se abren dos ventanas: el navegador con la app y el **Inspector**. Hacé el flujo a mano una vez: login, búsqueda del JPV, Actions, Change Case Type (si corresponde), Reassign, assignee, Submit y Close. El Inspector muestra una línea por cada acción, por ejemplo:

```python
page.get_by_role("button", name="Actions").click()
page.get_by_role("menuitem", name="Reassign").click()
```

Pasá esos valores a la sección `CONFIGURACIÓN` de `ui_automation.py` usando strings:

| Codegen | En el script |
|---|---|
| `get_by_role("textbox", name="Search")` | `"role=textbox[name='Search']"` |
| `get_by_role("button", name="Actions")` | `"role=button[name='Actions']"` |

Si no querés usar codegen, en el navegador: F12, clic derecho sobre el elemento, **Copy > Copy selector**.

Si la app es interna, necesitás estar en la red o en la VPN.

## 3. Configurar el script

Abrí `ui_automation.py` y completá la sección `CONFIGURACIÓN`:

**Búsqueda por JPV**
- `SEARCH_INPUT`: la search bar donde va el JPV.
- `SEARCH_BUTTON`: la lupa. Solo se usa si `SEARCH_WITH_ENTER` es `False`.
- `SEARCH_WITH_ENTER`: `True` confirma la búsqueda con Enter. `False` hace click en la lupa.

**Acciones sobre el caso**
- `ACTIONS_MENU`: el menú **Actions**.
- `REASSIGN_OPTION`: la opción **Reassign** dentro del menú.
- `OPERATOR_LABEL`: el label del dropdown **operator**, donde se elige el assignee. Por defecto `"operator"`.
- `SUBMIT_BUTTON` y `CLOSE_BUTTON`: los botones de cada paso.
- `CHANGE_CASE_TYPE_OPTION`: la opción **Change case type** dentro del menú Actions.
- `STEP_BY_STEP_OPTION`: el menuitem **Step By Step:** (se usa si **Reassign** no aparece en el menú).
- `MENU_CHECK_MS`: cuánto esperar a que aparezca una opción del menú **Actions** antes de decidir si se usa.
- `NEW_CASE_TYPE_LABEL`: el label del dropdown **Select New Case Type**.

**Archivo y navegador**
- `APP_URL`: la URL de la app. No la pongas en el script: usá `config_local.py` (ver abajo).
- `AUTO_SSO_CLICKS`: `True` hace los clicks de SSO (**Login with SSO** y **JPMorgan Chase (Default)**). `False` si preferís hacerlos a mano.
- `LOGIN_TIMEOUT_MS`: cuánto espera el script a que completes usuario, contraseña y MFA. Al terminar aparece la búsqueda y el script arranca solo.
- `XLSX_SHEET`: `None` usa la primera hoja. Si el archivo tiene varias, poné el nombre, por ejemplo `"Hoja1"`.
- `XLSX_HAS_HEADER`: `True` si la primera fila es de encabezado (se saltea).
- `COL_ASSIGNEE`, `COL_JPV`, `COL_TEAM` y `COL_CASE_TYPE`: letras de las columnas. Por defecto `A`, `B`, `C` y `D`.
- `CASE_TYPE_TRIGGER` y `NEW_CASE_TYPE`: si el tipo de caso es `CASE_TYPE_TRIGGER` (por defecto `"Miscellaneous"`), antes de reasignar se cambia a `NEW_CASE_TYPE`. `NEW_CASE_TYPE` es el **value** de la opción en el dropdown, no el texto visible.
- `TEAMS_ALLOWED`: conjunto de equipos que se procesan. Por defecto `{"AUTOCV", "FX ChaseTr", "iFX"}`. Las filas de otros equipos quedan como `OMITIDA`. La comparación distingue mayúsculas y minúsculas.
- `NAME_ALIASES`: nombres del XLSX que no coinciden con el dropdown, como `{"Nombre en XLSX": "Nombre en la app"}`.
- `BROWSER_CHANNEL`: `"chrome"` o `"msedge"` usa el navegador que ya tenés instalado. `None` usa el Chromium de Playwright.

**URL de la app (sin subirla al repo)**

1. Copiá `config_local.example.py` como `config_local.py`, en esta misma carpeta.
2. Completá `APP_URL` con la URL real.

`config_local.py` está en `.gitignore`, así que no se sube. El script lo usa automáticamente si existe.

**Checklist antes de la primera corrida:**
1. Todos los selectores tienen valores reales, no los de ejemplo.
2. `SEARCH_WITH_ENTER` coincide con cómo se confirma la búsqueda en la app.
3. El nombre de la columna A coincide exactamente con el texto de la opción en el dropdown **operator**.
4. `BROWSER_CHANNEL` está como querés.
5. `HEADLESS = False` mientras pruebas, para ver qué hace el script.

## 4. Preparar el XLSX

La columna A es el assignee, la columna B es el JPV, la columna C es el equipo y la columna D es el tipo de caso. Solo se procesan las filas cuyo equipo está en `TEAMS_ALLOWED`. El resto de las columnas no se usa, así que pueden quedar como estén.

| Assignee | JPV | Equipo | Tipo de caso | (otras columnas, se ignoran) |
|---|---|---|---|---|
| Alan David Garber | JPV2026113-000000 | AUTOCV | Miscellaneous | ... |
| Alan David Garber | JPV2026113-000001 | FX ChaseTr | Other | ... |

Armá tu XLSX con esa estructura. Los archivos `.xlsx` no se suben al repo. Las filas sin assignee o sin JPV se omiten y el script lo avisa en la terminal.

## 5. Ejecutar

Entrá a esta carpeta y corré el script desde ahí:

```bash
cd asignar-casos-empleados
```

Probá primero con un XLSX de prueba de una sola fila, con un caso que puedas revertir. Después:

```bash
python ui_automation.py datos.xlsx
```

1. Se abre el navegador en la app y el script hace los clicks de SSO.
2. Completá usuario, contraseña y MFA en el navegador.
3. Volvé a la terminal y presioná Enter.
4. El script procesa cada fila.

Al terminar:

- `resultados.csv`: el estado de cada fila: `OK`, `ERROR`, `OMITIDA` (equipo fuera de `TEAMS_ALLOWED`) o `NO PROCESADA` (la corrida se detuvo antes).
- `error_fila_N.png`: captura de la pantalla donde falló la fila N del XLSX.

`OK` indica que los pasos corrieron sin error. Confirmá en la app que la reasignación quedó guardada.

## Notas

- Los selectores que vienen en el script son placeholders. Hay que reemplazarlos antes de usarlo.
- Si un assignee de la columna A no existe en el dropdown **operator**, esa fila queda en `ERROR` y el script sigue con la siguiente.
- Si una fila falla, el script guarda una captura, recarga la página y vuelve a la search bar. Si la recuperación funciona, sigue con la siguiente fila.
- Si no logra volver a la search bar, se detiene y marca las filas restantes como `NO PROCESADA` en `resultados.csv`. Revisá la pantalla antes de volver a correrlo.
- El mensaje de error dice en qué paso falló, por ejemplo `paso 'Actions'`.
- El archivo de entrada tiene que ser `.xlsx`. Otros formatos (CSV, XLS) no se aceptan.
