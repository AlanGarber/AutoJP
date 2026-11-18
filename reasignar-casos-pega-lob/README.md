# Reasignar casos en Pega según LOB

Reasigna casos de Pega a la basket que corresponde a su LOB, a partir de un XLSX. Por cada fila:

1. Busca el JPV (columna C, **ID**) en la search bar de la basket. La basket es un popup.
2. Hace click en el JPV. El caso se abre en la ventana principal de Pega, dentro de iframes `PegaGadget*`.
3. Hace click en **Reassign**.
4. Elige la basket de destino según la LOB (columna A, **Comments**): `Corp&FI` → `FX12`, `Retail` → `FX13`.
5. Confirma la reasignación.

Todo está en un solo archivo: `ui_automation.py`.

> **Estado: borrador.** Los selectores marcados como `PLACEHOLDER` en el script todavía no están confirmados, y el diálogo de Reassign (paso 4 y 5) no está mapeado. Por eso el script arranca con `DRY_RUN = True`: navega, abre Reassign y cancela sin reasignar.

## 1. Instalación

Necesitás Python 3.9 o superior.

```bash
pip install playwright openpyxl
```

Si usás Edge o Chrome ya instalados, no hace falta `playwright install`. El script usa `BROWSER_CHANNEL = "msedge"` por defecto. Si preferís el Chromium de Playwright, poné `BROWSER_CHANNEL = None` y ejecutá `playwright install chromium`.

## 2. Sacar los selectores con codegen

```bash
playwright codegen --channel msedge https://TU-APP-PEGA
```

Hacé el flujo a mano una vez: login, abrir la basket, buscar un JPV, clickear el JPV, Reassign, elegir destino y cancelar (no confirmes sobre un caso real). El Inspector muestra una línea por acción. Pasá esos valores a la sección `CONFIGURACIÓN` del script.

Si tenés que copiar el código a mano, tené en cuenta que el codegen también graba lo que tipeás en el login. No dejes contraseñas en el archivo.

## 3. Configurar el script

**URL de la app (sin subirla al repo)**

1. Copiá `config_local.example.py` como `config_local.py`, en esta misma carpeta.
2. Completá `APP_URL` con la URL real de Pega.

`config_local.py` está en `.gitignore`, así que no se sube.

**Selectores (sección `CONFIGURACIÓN` de `ui_automation.py`)**
- `BASKET_SEARCH_INPUT`: search bar de la basket.
- `CASE_LINK`: link del JPV en la grilla. `{jpv}` se reemplaza por el JPV de cada fila.
- `REASSIGN_BUTTON`: botón Reassign dentro del caso. Confirmá si es un botón o un menuitem.
- `CANCEL_REASSIGN`: botón para cerrar el diálogo en modo simulación.

**Excel y mapeo**
- `COL_LOB` y `COL_JPV`: letras de las columnas. Por defecto `A` y `C`.
- `MAPA_LOB`: texto exacto de la LOB → basket de destino.
- `XLSX_HAS_HEADER`: `True` si la primera fila es encabezado.

**Ejecución**
- `DRY_RUN`: `True` simula. `False` reasigna de verdad (todavía no implementado, ver abajo).
- `BROWSER_CHANNEL`: `"msedge"`, `"chrome"` o `None`.

## 4. Preparar el XLSX

| Comments (A) | … | ID (C) | … |
|---|---|---|---|
| Corp&FI | … | JPV2026118-000000 | … |
| Retail | … | JPV2026118-000001 | … |

Las demás columnas se ignoran. El archivo tiene que ser `.xlsx`. No subas archivos con datos al repo.

## 5. Ejecutar

Probá primero con un XLSX de una sola fila.

```bash
cd reasignar-casos-pega-lob
python ui_automation.py datos.xlsx
```

1. Se abre el navegador en Pega. Hacé el login (usuario, contraseña y MFA).
2. Abrí la basket como popup desde la ventana de Pega.
3. Volvé a la terminal y presioná Enter.
4. El script procesa cada fila.

Al terminar:
- `resultados.csv`: estado de cada fila. `SIMULADO` (DRY_RUN), `OK`, `ERROR`, `OMITIDA` (LOB sin basket o JPV vacío) o `NO PROCESADA` (la basket se cerró antes).
- `error_fila_N.png`: captura de la ventana principal en la fila N, si falló.

## Pendiente antes de usar `DRY_RUN = False`

- Confirmar los selectores `PLACEHOLDER`.
- Completar `completar_reasignacion()`: el campo de basket destino y el botón de confirmación del diálogo de Reassign.
- Definir qué hacer si el JPV no aparece en la basket (hoy queda como `ERROR` y el script sigue).
