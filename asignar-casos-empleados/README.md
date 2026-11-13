# Asignar casos a cada empleado según CSV

Asigna casos a empleados en una UI web a partir de un CSV. Por cada fila:

1. Escribe el caso (columna 1) en la search bar y presiona Enter.
2. Abre el dropdown y elige una opción.
3. Presiona Tab, escribe el empleado (columna 2) y presiona Enter.
4. Hace click en el botón de asignación.

Todo está en un solo archivo: `ui_automation.py`.

## 1. Instalación

Necesitás Python 3.9 o superior.

```bash
pip install playwright
playwright install chromium
```

## 2. Sacar los selectores con codegen

Playwright tiene un grabador que escribe el código por vos.

```bash
playwright codegen https://TU-APP-INTERNA
```

Se abren dos ventanas: el navegador con la app y el **Inspector**. Hacé el flujo a mano una vez (search bar, dropdown, opción, Tab, empleado, botón). El Inspector muestra una línea por cada acción, por ejemplo:

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

Abrí `ui_automation.py` y completá:

- `APP_URL`: la URL de la app.
- `SEARCH_INPUT`: la search bar donde se busca el caso.
- `DROPDOWN`: el elemento que abre el dropdown.
- `DROPDOWN_OPTION`: el texto de la opción a elegir.
- `SUBMIT_BUTTON`: el botón de asignación.

Si el CSV usa `;` como separador (común con Excel en configuración regional de Argentina), cambiá `CSV_DELIMITER = ";"`.

## 4. Preparar el CSV

La primera columna es el caso y la segunda el empleado. El script espera una fila de encabezado (`CSV_HAS_HEADER = True`).

```csv
caso,empleado
JPV2026113-000000,Alan David Garber
```

Guardalo en UTF-8 para que no se rompan las tildes y la ñ. Un ejemplo listo está en `datos_ejemplo.csv`.

## 5. Ejecutar

Entrá a esta carpeta y corré el script desde ahí:

```bash
cd asignar-casos-empleados
```

Probá primero con `datos_ejemplo.csv` (2 filas) o con 2 o 3 filas de tu archivo. Después:

```bash
python ui_automation.py datos.csv
```

1. Se abre el navegador en la app.
2. Logueate (SSO si corresponde).
3. Volvé a la terminal y presioná Enter.
4. El script procesa cada fila.

Al terminar:

- `resultados.csv`: el estado (`OK` o `ERROR`) de cada caso.
- `error_fila_N.png`: captura de la pantalla donde falló la fila N.

## Notas

- Los selectores que vienen en el script son placeholders. Hay que reemplazarlos antes de usarlo.
- Si una fila falla, el script sigue con la siguiente.
