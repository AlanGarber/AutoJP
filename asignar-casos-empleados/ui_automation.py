#!/usr/bin/env python3
"""
Asigna casos a cada empleado en Pega a partir de un XLSX (una pasada por fila).

Por cada fila:
  1. Busca el JPV (columna B) en la search bar y confirma la búsqueda.
  2. Abre Actions y elige Reassign.
  3. Escribe el assignee (columna A) en el campo "label operator".
  4. Hace click en Submit, Close y Save.

Al inicio, el script hace los clicks de SSO (si AUTO_SSO_CLICKS es True).
Usuario, contraseña y MFA los completás vos antes de que empiece la carga.
Las demás columnas del archivo se ignoran.

Instalación:
    pip install playwright openpyxl
    playwright install chromium

Uso:
    python ui_automation.py datos.xlsx

Genera resultados.csv con el estado de cada fila y una captura
(error_fila_N.png) por cada fila que falle.
"""

import csv
import logging
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from playwright.sync_api import Page, sync_playwright

# =====================================================================
# CONFIGURACIÓN: completar con los selectores reales de la app.
# Tip: `playwright codegen <URL>` graba los clicks y devuelve selectores.
# Los valores de ejemplo son placeholders.
# =====================================================================
APP_URL = "https://TU-APP-INTERNA"

# --- Búsqueda por JPV ---
SEARCH_INPUT = "input[placeholder='Search']"       # search bar donde va el JPV
SEARCH_BUTTON = "role=button[name='Search']"        # lupa (se usa si SEARCH_WITH_ENTER es False)
SEARCH_WITH_ENTER = True                            # True: Enter confirma la búsqueda; False: click en la lupa

# --- Acciones sobre el caso ---
ACTIONS_MENU = "role=button[name='Actions']"        # PegaGadget Actions
REASSIGN_OPTION = "role=menuitem[name='Reassign']"  # PegaGadget Reassign
OPERATOR_INPUT = "input[aria-label='Operator']"     # campo "label operator", donde va el assignee
SUBMIT_BUTTON = "role=button[name='Submit']"        # PegaGadget Submit
CLOSE_BUTTON = "role=button[name='Close']"          # PegaGadget Close
SAVE_BUTTON = "role=button[name='Save']"            # PegaGadget Save

XLSX_SHEET = None          # None = primera hoja; o el nombre de la hoja, por ejemplo "Hoja1"
XLSX_HAS_HEADER = True     # True: la primera fila es encabezado y se saltea
COL_ASSIGNEE = "A"         # columna con el assignee
COL_JPV = "B"              # columna con el identificador JPV

BROWSER_CHANNEL = "chrome"  # "chrome" o "msedge": usa el navegador ya instalado; None: Chromium de Playwright
HEADLESS = False        # False: ves la automatización corriendo
AUTO_SSO_CLICKS = True  # True: el script hace click en "Login with SSO" y en "JPMorgan Chase (Default)"
LOGIN_PAUSE = True      # True: pausa para completar usuario, contraseña y MFA antes de empezar
TYPING_DELAY_MS = 40    # tipeo humano; ayuda con UIs que reaccionan a cada tecla
STEP_PAUSE_MS = 300     # pausa breve entre pasos

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s %(message)s",
)
log = logging.getLogger("ui_automation")


def to_text(value) -> str:
    """Convierte una celda de Excel a texto. Los enteros no quedan con '.0'."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def read_rows(path: Path):
    """Devuelve [(nro_fila, assignee, jpv)] para cada fila válida del XLSX."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[XLSX_SHEET] if XLSX_SHEET else wb.active
        idx_assignee = column_index_from_string(COL_ASSIGNEE) - 1
        idx_jpv = column_index_from_string(COL_JPV) - 1
        first_row = 2 if XLSX_HAS_HEADER else 1

        rows = []
        for row_no, cells in enumerate(
            ws.iter_rows(min_row=first_row, values_only=True), start=first_row
        ):
            assignee = to_text(cells[idx_assignee]) if idx_assignee < len(cells) else ""
            jpv = to_text(cells[idx_jpv]) if idx_jpv < len(cells) else ""
            if not assignee or not jpv:
                if assignee or jpv:  # fila incompleta: se avisa, no se procesa
                    log.warning("Fila %d omitida (falta assignee o JPV)", row_no)
                continue
            rows.append((row_no, assignee, jpv))
        return rows
    finally:
        wb.close()


def login_sso(page: Page) -> None:
    """Clicks de SSO. Usuario, contraseña y MFA los completa la persona."""
    page.get_by_role("link", name="Login with SSO", exact=True).click()
    page.get_by_role("button", name="JPMorgan Chase (Default)").click()
    page.wait_for_timeout(STEP_PAUSE_MS)


def process_row(page: Page, assignee: str, jpv: str) -> None:
    # 1) Buscar el JPV
    search = page.locator(SEARCH_INPUT)
    search.click()
    search.fill("")
    search.press_sequentially(jpv, delay=TYPING_DELAY_MS)
    if SEARCH_WITH_ENTER:
        search.press("Enter")
    else:
        page.locator(SEARCH_BUTTON).click()
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 2) Actions > Reassign
    page.locator(ACTIONS_MENU).click()
    page.locator(REASSIGN_OPTION).click()
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 3) Escribir el assignee en "label operator"
    operator = page.locator(OPERATOR_INPUT)
    operator.click()
    operator.fill("")
    operator.press_sequentially(assignee, delay=TYPING_DELAY_MS)
    # Si la app muestra sugerencias y hay que elegir una, agregar el click acá.
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 4) Submit, Close y Save
    page.locator(SUBMIT_BUTTON).click()
    page.wait_for_timeout(STEP_PAUSE_MS)
    page.locator(CLOSE_BUTTON).click()
    page.wait_for_timeout(STEP_PAUSE_MS)
    page.locator(SAVE_BUTTON).click()
    page.wait_for_timeout(STEP_PAUSE_MS)


def main() -> None:
    if len(sys.argv) != 2:
        sys.exit("Uso: python ui_automation.py archivo.xlsx")

    xlsx_path = Path(sys.argv[1])
    if xlsx_path.suffix.lower() != ".xlsx":
        sys.exit("El archivo de entrada tiene que ser .xlsx")

    rows = read_rows(xlsx_path)
    log.info("%d filas para procesar", len(rows))
    if not rows:
        return

    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(channel=BROWSER_CHANNEL, headless=HEADLESS)
        page = browser.new_page()
        page.goto(APP_URL)

        if AUTO_SSO_CLICKS:
            login_sso(page)

        if LOGIN_PAUSE:
            input("Completá usuario, contraseña y MFA en el navegador y presioná Enter para empezar... ")

        for row_no, assignee, jpv in rows:
            try:
                process_row(page, assignee, jpv)
                log.info("Fila %d OK (%s | %s)", row_no, jpv, assignee)
                results.append([row_no, assignee, jpv, "OK", ""])
            except Exception as exc:  # noqa: BLE001
                text = str(exc).strip()
                msg = text.splitlines()[0] if text else type(exc).__name__
                log.error("Fila %d ERROR: %s", row_no, msg)
                try:
                    page.screenshot(path=f"error_fila_{row_no}.png", full_page=True)
                    page.keyboard.press("Escape")  # cerrar menús o dialogs colgados
                except Exception:  # la página pudo quedar en mal estado
                    pass
                results.append([row_no, assignee, jpv, "ERROR", msg])
        browser.close()

    out = Path("resultados.csv")
    with out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["fila", "assignee", "jpv", "estado", "detalle"])
        writer.writerows(results)

    ok = sum(r[3] == "OK" for r in results)
    log.info("Terminado: %d OK, %d con error. Detalle en %s", ok, len(results) - ok, out.resolve())


if __name__ == "__main__":
    main()
