# Reasignar casos en Pega según LOB

Reasigna casos de Pega a la work queue que corresponde a su LOB, a partir de un XLSX. Por cada fila:

1. Hace click en el JPV (columna C, **ID**) en la basket. La basket es un popup.
2. El caso se abre en la ventana principal de Pega. Pestaña **Work**: si el case owner no es `TSFX1213`, la fila se saltea.
3. **Actions** > **Reassign**.
4. Elige la work queue según la LOB (columna A, **Comments**) en el dropdown **Work Queue**: `Corp&FI` → `WDPCTSTWRKFX12`, `Retail` → `WDPCTSTWRKFX13`.
5. **Submit** y **Close**. En modo simulación, **Close** y **Discard** (no se guarda nada).

Todo está en un solo archivo: `ui_automation.py`.

> **Estado: borrador.** Los valores marcados `PLACEHOLDER` en el script todavía no están confirmados. Por eso el script arranca con `DRY_RUN = True`.

## 1. Instalación

Necesitás Python 3.9 o superior.

```bash
pip install playwright openpyxl
```

El script usa `BROWSER_CHANNEL = "msedge"` por defecto, así que no hace falta `playwright install`. Si preferís el Chromium de Playwright, poné `BROWSER_CHANNEL = None` y ejecutá `playwright install chromium`.

## 2. Sacar los selectores con codegen

```bash
playwright codegen --channel msedge --target python -o pasada_pega.py https://TU-APP-PEGA
```

Hacé el flujo a mano una vez: login, abrir la basket desde **View Queue for**, clickear el JPV, Actions > Reassign, elegir la Work Queue y cerrar con Close y Discard (no confirmes sobre un caso real). El Inspector muestra una línea por acción.

Codegen graba lo que tipeás en el login. No dejes contraseñas en el archivo antes de compartirlo.

## 3. Configurar el script

**URL de la app (sin subirla al repo)**

1. Copiá `config_local.example.py` como `config_local.py`, en esta misma carpeta.
2. Completá `APP_URL` con la URL real de Pega.

`config_local.py` está en `.gitignore`, así que no se sube.

**Excel y mapeo**
- `COL_LOB` y `COL_JPV`: letras de las columnas. Por defecto `A` y `C`.
- `MAPA_LOB`: texto exacto de la LOB → value de la work queue en el dropdown.
- `XLSX_HAS_HEADER`: `True` si la primera fila es encabezado.

**Selectores (sección `CONFIGURACIÓN` de `ui_automation.py`)**
- `JPV_LINK_NTH`: el índice del link del JPV en la basket (`2`, según codegen).
- `OWNER_EXPECTED`: case owner que tiene que tener el caso para reasignarse.
- `WORK_QUEUE_SELECTOR`: dropdown de la work queue destino.

**Ejecución**
- `DRY_RUN`: `True` simula. `False` reasigna de verdad.
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
2. Abrí la basket desde **View Queue for** (popup).
3. Volvé a la terminal y presioná Enter.
4. El script procesa cada fila.

Al terminar:
- `resultados.csv`: estado de cada fila. `SIMULADO` (DRY_RUN), `OK`, `ERROR`, `OMITIDA` (case owner distinto de `TSFX1213`, LOB sin work queue o JPV vacío) o `NO PROCESADA` (la basket se cerró antes).
- `error_fila_N.png`: captura de la ventana principal en la fila N, si falló.

## Pendiente antes de usar `DRY_RUN = False`

- Cambiar la validación del case owner por igualdad exacta, con el locator del campo **Case owner** en la pestaña Work. Hoy busca `TSFX1213` como palabra completa dentro de todo el panel.
- Confirmar los selectores `PLACEHOLDER` contra la app real.
