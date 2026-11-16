# Asignar casos a cada empleado según XLSX

Asigna casos a empleados en una UI web a partir de un archivo XLSX. Por cada fila:

1. Escribe el assignee (columna A) en la search bar y presiona Enter.
2. Abre el dropdown y elige una opción.
3. Presiona Tab, escribe el JPV (columna B) y presiona Enter.
4. Hace click en el botón de asignación.

Las demás columnas del archivo se ignoran. Todo está en un solo archivo: `ui_automation.py`.

## 1. Instalación

Necesitás Python 3.9 o superior.

```bash
pip install playwright openpyxl
playwright install chromium
```

Si `playwright install chromium` falla con `ETIMEDOUT` o un error de descarga, la red corporativa está bloqueando la descarga. Dos opciones:

- **Usar Chrome o Edge ya instalados (recomendado):** en `ui_automation.py` dejá `BROWSER_CHANNEL = "chrome"` (o `"msedge"`). No hace falta descargar nada. Para codegen, usá la opción `--channel chrome` (ver paso 2).
- **Configurar el proxy y reintentar:** en PowerShell, `$env:HTTPS_PROXY="http://PROXY:PUERTO"` y después `playwright install chromium`. La dirección del proxy la pide IT.

## 2. Sacar los selectores con codegen

Playwright tiene un grabador que escribe el código por vos.

```bash
playwright codegen --channel chrome https://TU-APP-INTERNA
```

Si no tenés Chrome instalado, quitá `--channel chrome`. Eso usa el Chromium de Playwright y requiere el paso de descarga de la sección 1.

Se abren dos ventanas: el navegador con la app y el **Inspector**. Hacé el flujo a mano una vez: assignee en la search bar, dropdown, opción, Tab, JPV y botón. El Inspector muestra una línea por cada acción, por ejemplo:

```python
page.get_by_placeholder("Search").click()
page.get_by_role("button", name="Submit").click()
```

Pasá esos valores a la sección `CONFIGURACIÓN` de `ui_automation.py` usando strings:

| Codegen | En el script |
|---|---|
| `get_by_placeholder("Search")` | `"input[placeholder='Search']"` |
| `get_by_role("button", name="Submit")` | `"role=button[name='Submit']"` |

Si no querés usar codegen, en Chrome: F12, clic derecho sobre el elemento, **Copy > Copy selector**.

Si la app es interna, necesitás estar en la red o en la VPN.

## 3. Configurar el script

Abrí `ui_automation.py` y completá la sección `CONFIGURACIÓN`:

- `APP_URL`: la URL de la app.
- `SEARCH_INPUT`: la search bar donde se busca el assignee.
- `DROPDOWN`: el elemento que abre el dropdown.
- `DROPDOWN_OPTION`: el texto exacto de la opción a elegir.
- `SUBMIT_BUTTON`: el botón de asignación.

Y si hace falta, ajustá las columnas del XLSX y el navegador:

- `XLSX_SHEET`: `None` usa la primera hoja. Si el archivo tiene varias, poné el nombre de la hoja, por ejemplo `"Hoja1"`.
- `XLSX_HAS_HEADER`: `True` si la primera fila es de encabezado (se saltea).
- `COL_ASSIGNEE` y `COL_JPV`: letras de las columnas. Por defecto `A` y `B`.
- `BROWSER_CHANNEL`: `"chrome"` o `"msedge"` usa el navegador que ya tenés instalado. `None` usa el Chromium de Playwright, que hay que descargar con `playwright install chromium`.

**Checklist antes de la primera corrida:**
1. `APP_URL`, `SEARCH_INPUT`, `DROPDOWN`, `DROPDOWN_OPTION` y `SUBMIT_BUTTON` tienen valores reales, no los de ejemplo.
2. `DROPDOWN_OPTION` coincide con el texto que muestra la app.
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

1. Se abre el navegador en la app.
2. Logueate (SSO si corresponde).
3. Volvé a la terminal y presioná Enter.
4. El script procesa cada fila.

Al terminar:

- `resultados.csv`: el estado (`OK` o `ERROR`) de cada fila.
- `error_fila_N.png`: captura de la pantalla donde falló la fila N del XLSX.

## Notas

- Los selectores que vienen en el script son placeholders. Hay que reemplazarlos antes de usarlo.
- `DROPDOWN_OPTION` es fijo: el script siempre elige la misma opción. Si la opción cambia según la fila, hay que modificar el script.
- Si una fila falla, el script sigue con la siguiente.
- El archivo de entrada tiene que ser `.xlsx`. Otros formatos (CSV, XLS) no se aceptan.
