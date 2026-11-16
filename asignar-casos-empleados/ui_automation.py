#!/usr/bin/env python3
"""
Asigna casos a cada empleado en una UI web a partir de un XLSX (una pasada por fila).

Por cada fila:
  1. Escribe el assignee (columna A) en la search bar y presiona Enter.
  2. Abre el dropdown y elige una opción.
  3. Presiona Tab, escribe el JPV (columna B) y presiona Enter.
  4. Hace click en el botón de asignación.

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
# =====================================================================
APP_URL = "https://TU-APP-INTERNA"

SEARCH_INPUT = "input[placeholder='Search']"   # search bar
DROPDOWN = "#dropdown-trigger"                 # elemento que abre el dropdown
DROPDOWN_OPTION = "Opción A"                   # texto visible de la opción a elegir
SUBMIT_BUTTON = "button:has-text('Submit')"    # botón final de cada fila

XLSX_SHEET = None          # None = primera hoja; o el nombre de la hoja, por ejemplo "Hoja1"
XLSX_HAS_HEADER = True     # True: la primera fila es encabezado y se saltea
COL_ASSIGNEE = "A"         # columna con el assignee
COL_JPV = "B"              # columna con el identificador JPV

HEADLESS = False        # False: ves la automatización corriendo
LOGIN_PAUSE = True      # True: pausa para loguearte (SSO) antes de empezar
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


def process_row(page: Page, assignee: str, jpv: str) -> None:
    # 1) Search bar: limpiar, escribir el assignee y Enter
    search = page.locator(SEARCH_INPUT)
    search.click()
    search.fill("")
    search.press_sequentially(assignee, delay=TYPING_DELAY_MS)
    search.press("Enter")
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 2) Dropdown: abrir y elegir la opción
    page.locator(DROPDOWN).click()
    page.get_by_text(DROPDOWN_OPTION, exact=True).click()
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 3) Tab al siguiente campo: reemplazar por el JPV y Enter
    page.keyboard.press("Tab")
    page.keyboard.press("ControlOrMeta+A")
    page.keyboard.type(jpv, delay=TYPING_DELAY_MS)
    page.keyboard.press("Enter")
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 4) Botón de asignación
    page.locator(SUBMIT_BUTTON).click()
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
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        page.goto(APP_URL)

        if LOGIN_PAUSE:
            input("Logueate en la app y presioná Enter para empezar... ")

        for row_no, assignee, jpv in rows:
            try:
                process_row(page, assignee, jpv)
                log.info("Fila %d OK (%s | %s)", row_no, assignee, jpv)
                results.append([row_no, assignee, jpv, "OK", ""])
            except Exception as exc:  # noqa: BLE001
                text = str(exc).strip()
                msg = text.splitlines()[0] if text else type(exc).__name__
                log.error("Fila %d ERROR: %s", row_no, msg)
                try:
                    page.screenshot(path=f"error_fila_{row_no}.png", full_page=True)
                    page.keyboard.press("Escape")  # cerrar dropdowns colgados
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
