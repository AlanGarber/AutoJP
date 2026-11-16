# Asignar casos a cada empleado según XLSX

Reasigna casos a empleados en Pega a partir de un archivo XLSX. Por cada fila:

1. Busca el JPV (columna B) en la search bar y confirma la búsqueda.
2. Abre **Actions** y elige **Reassign**.
3. Escribe el assignee (columna A) en el campo **label operator**.
4. Hace click en **Submit**, **Close** y **Save**.

Al inicio, el script hace los clicks de SSO (**Login with SSO** y **JPMorgan Chase (Default)**). Usuario, contraseña y MFA los completás vos. Las demás columnas del archivo se ignoran. Todo está en un solo archivo: `ui_automation.py`.

## 1. Instalación

Necesitás Python 3.9 o superior.

```bash
pip install playwright openpyxl
playwright install chromium
```

Si `playwright install chromium` falla con `ETIMEDOUT` o un error de descarga, la red corporativa está bloqueando la descarga. Dos opciones:

- **Usar Chrome o Edge ya instalados (recomendado):** en `ui_automation.py` poné `BROWSER_CHANNEL = "chrome"` o `"msedge"`. No hace falta descargar nada. Para codegen, usá `--channel chrome` o `--channel msedge` (ver paso 2).
- **Configurar el proxy y reintentar:** en PowerShell, `$env:HTTPS_PROXY="http://PROXY:PUERTO"` y después `playwright install chromium`. La dirección del proxy la pide IT.

## 2. Sacar los selectores con codegen

Playwright tiene un grabador que escribe el código por vos.

```bash
playwright codegen --channel msedge https://TU-APP-INTERNA
```

Si no usás Edge, cambiá `msedge` por `chrome`. Si no tenés ninguno de los dos, quitá `--channel` y usá el Chromium de Playwright, que requiere el paso de descarga de la sección 1.

Se abren dos ventanas: el navegador con la app y el **Inspector**. Hacé el flujo a mano una vez: login, búsqueda del JPV, Actions, Reassign, assignee, Submit, Close y Save. El Inspector muestra una línea por cada acción, por ejemplo:

```python
page.get_by_role("button", name="Actions").click()
page.get_by_role("menuitem", name="Reassign").click()
```

Pasá esos valores a la sección `CONFIGURACIÓN` de `ui_automation.py` usando strings:

| Codegen | En el script |
|---|---|
| `get_by_placeholder("Search")` | `"input[placeholder='Search']"` |
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
- `OPERATOR_INPUT`: el campo **label operator**, donde va el assignee.
- `SUBMIT_BUTTON`, `CLOSE_BUTTON` y `SAVE_BUTTON`: los botones de cada paso.

**Archivo y navegador**
- `APP_URL`: la URL de la app.
- `AUTO_SSO_CLICKS`: `True` hace los clicks de SSO (**Login with SSO** y **JPMorgan Chase (Default)**). `False` si preferís hacerlos a mano.
- `LOGIN_PAUSE`: `True` pausa después del login para que completes usuario, contraseña y MFA.
- `XLSX_SHEET`: `None` usa la primera hoja. Si el archivo tiene varias, poné el nombre, por ejemplo `"Hoja1"`.
- `XLSX_HAS_HEADER`: `True` si la primera fila es de encabezado (se saltea).
- `COL_ASSIGNEE` y `COL_JPV`: letras de las columnas. Por defecto `A` y `B`.
- `BROWSER_CHANNEL`: `"chrome"` o `"msedge"` usa el navegador que ya tenés instalado. `None` usa el Chromium de Playwright.

**Checklist antes de la primera corrida:**
1. Todos los selectores tienen valores reales, no los de ejemplo.
2. `SEARCH_WITH_ENTER` coincide con cómo se confirma la búsqueda en la app.
3. `BROWSER_CHANNEL` está como querés.
4. `HEADLESS = False` mientras pruebas, para ver qué hace el script.

## 4. Preparar el XLSX

La columna A es el assignee y la columna B es el JPV. El resto de las columnas no se usa, así que pueden quedar como estén.

| Assignee | JPV | (otras columnas, se ignoran) |
|---|---|---|
| Alan David Garber | JPV2026113-000000 | ... |
| Alan David Garber | JPV2026113-000001 | ... |

Un ejemplo listo está en `datos_ejemplo.xlsx`. Las filas sin assignee o sin JPV se omiten y el script lo avisa en la terminal.

## 5. Ejecutar

Entrá a esta carpeta y corré el script desde ahí:

```bash
cd asignar-casos-empleados
```

Probá primero con `datos_ejemplo.xlsx` (2 filas) o con 2 o 3 filas de tu archivo. Después:

```bash
python ui_automation.py datos.xlsx
```

1. Se abre el navegador en la app y el script hace los clicks de SSO.
2. Completá usuario, contraseña y MFA en el navegador.
3. Volvé a la terminal y presioná Enter.
4. El script procesa cada fila.

Al terminar:

- `resultados.csv`: el estado (`OK` o `ERROR`) de cada fila.
- `error_fila_N.png`: captura de la pantalla donde falló la fila N del XLSX.

`OK` indica que los pasos corrieron sin error. Confirmá en la app que la reasignación quedó guardada.

## Notas

- Los selectores que vienen en el script son placeholders. Hay que reemplazarlos antes de usarlo.
- Si la app muestra sugerencias al escribir en el campo del assignee, hay que agregar el click en la opción correcta en `process_row`.
- Si una fila falla, el script sigue con la siguiente.
- El archivo de entrada tiene que ser `.xlsx`. Otros formatos (CSV, XLS) no se aceptan.
