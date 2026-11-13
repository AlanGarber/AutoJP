#!/usr/bin/env python3
"""
Automatiza una UI web a partir de un CSV (una pasada completa por fila).

Por cada fila:
  1. Escribe la columna 1 en la search bar y presiona Enter.
  2. Abre el dropdown y elige una opción.
  3. Presiona Tab, escribe la columna 2 y presiona Enter.
  4. Hace click en el botón.

Instalación:
    pip install playwright
    playwright install chromium

Uso:
    python ui_automation.py datos.csv

Genera resultados.csv con el estado de cada fila y una captura
(error_fila_N.png) por cada fila que falle.
"""

import csv
import logging
import sys
from pathlib import Path

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

CSV_DELIMITER = ","
CSV_HAS_HEADER = True
CSV_ENCODING = "utf-8-sig"

HEADLESS = False        # False: ves la automatización corriendo
LOGIN_PAUSE = True      # True: pausa para loguearte (SSO) antes de empezar
TYPING_DELAY_MS = 40    # tipeo humano; ayuda con UIs que reaccionan a cada tecla
STEP_PAUSE_MS = 300     # pausa breve entre pasos

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s %(message)s",
)
log = logging.getLogger("ui_automation")


def read_rows(path: Path):
    """Devuelve (nro_linea, col1, col2) por cada fila válida del CSV."""
    with path.open(newline="", encoding=CSV_ENCODING) as f:
        reader = csv.reader(f, delimiter=CSV_DELIMITER)
        if CSV_HAS_HEADER:
            next(reader, None)
        start = 2 if CSV_HAS_HEADER else 1
        for line_no, row in enumerate(reader, start=start):
            if len(row) < 2 or not row[0].strip():
                log.warning("Fila %d omitida (vacía o con menos de 2 columnas)", line_no)
                continue
            yield line_no, row[0].strip(), row[1].strip()


def process_row(page: Page, col1: str, col2: str) -> None:
    # 1) Search bar: limpiar, escribir col1 y Enter
    search = page.locator(SEARCH_INPUT)
    search.click()
    search.fill("")
    search.press_sequentially(col1, delay=TYPING_DELAY_MS)
    search.press("Enter")
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 2) Dropdown: abrir y elegir la opción
    page.locator(DROPDOWN).click()
    page.get_by_text(DROPDOWN_OPTION, exact=True).click()
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 3) Tab al siguiente campo: reemplazar por col2 y Enter
    page.keyboard.press("Tab")
    page.keyboard.press("ControlOrMeta+A")
    page.keyboard.type(col2, delay=TYPING_DELAY_MS)
    page.keyboard.press("Enter")
    page.wait_for_timeout(STEP_PAUSE_MS)

    # 4) Botón de la UI
    page.locator(SUBMIT_BUTTON).click()
    page.wait_for_timeout(STEP_PAUSE_MS)


def main() -> None:
    if len(sys.argv) != 2:
        sys.exit("Uso: python ui_automation.py archivo.csv")

    csv_path = Path(sys.argv[1])
    rows = list(read_rows(csv_path))
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

        for line_no, col1, col2 in rows:
            try:
                process_row(page, col1, col2)
                log.info("Fila %d OK (%s | %s)", line_no, col1, col2)
                results.append([line_no, col1, col2, "OK", ""])
            except Exception as exc:  # noqa: BLE001
                text = str(exc).strip()
                msg = text.splitlines()[0] if text else type(exc).__name__
                log.error("Fila %d ERROR: %s", line_no, msg)
                try:
                    page.screenshot(path=f"error_fila_{line_no}.png", full_page=True)
                    page.keyboard.press("Escape")  # cerrar dropdowns colgados
                except Exception:  # la página pudo quedar en mal estado
                    pass
                results.append([line_no, col1, col2, "ERROR", msg])
        browser.close()

    out = Path("resultados.csv")
    with out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["linea", "col1", "col2", "estado", "detalle"])
        writer.writerows(results)

    ok = sum(r[3] == "OK" for r in results)
    log.info("Terminado: %d OK, %d con error. Detalle en %s", ok, len(results) - ok, out.resolve())


if __name__ == "__main__":
    main()
