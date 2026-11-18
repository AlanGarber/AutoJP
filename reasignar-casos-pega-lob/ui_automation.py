#!/usr/bin/env python3
"""
Reasigna casos de Pega a la basket que corresponde a su LOB, a partir de un XLSX (una pasada por fila).

Por cada fila:
  1. Busca el JPV (columna C, "ID") en la search bar de la basket. La basket es un popup.
  2. Hace click en el JPV: el caso se abre en la ventana principal de Pega, dentro de iframes PegaGadget*.
  3. Hace click en Reassign.
  4. Elige la basket de destino según la LOB (columna A, "Comments"): Corp&FI -> FX12, Retail -> FX13.
  5. Confirma la reasignación. (Pendiente: el diálogo de Reassign todavía no está mapeado.)

Antes de la primera fila, vos hacés el login en la ventana principal y abrís la basket
(popup). Después volvés a la terminal y presionás Enter.

Uso:
    python ui_automation.py datos.xlsx

Por defecto corre con DRY_RUN = True: abre el caso, hace click en Reassign y cancela sin confirmar.
Genera resultados.csv y una captura (error_fila_N.png) por cada fila que falle.
"""

import csv
import logging
import os
import sys
import time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from playwright.sync_api import BrowserContext, Page, sync_playwright

# =====================================================================
# CONFIGURACIÓN: completar con los selectores reales de Pega.
# Tip: `playwright codegen --channel msedge <URL>` graba los clicks y devuelve selectores.
# Los valores marcados PLACEHOLDER todavía no se confirmaron.
# =====================================================================
APP_URL = "https://TU-APP-PEGA"
try:
    # URL real, en config_local.py (ver config_local.example.py). No se sube al repo.
    from config_local import APP_URL
except ImportError:
    pass

# --- Excel ---
XLSX_SHEET = None              # None = primera hoja; o el nombre de la hoja, por ejemplo "Hoja1"
XLSX_HAS_HEADER = True         # True: la primera fila es encabezado y se saltea
COL_LOB = "A"                  # columna "Comments": LOB
COL_JPV = "C"                  # columna "ID": JPV (ID del caso)
MAPA_LOB = {                   # LOB (texto exacto del XLSX) -> basket de destino
    "Corp&FI": "FX12",
    "Retail": "FX13",
}

# --- Basket (popup) ---
BASKET_SEARCH_INPUT = "role=textbox[name='Search'] >> visible=true"  # PLACEHOLDER: search bar de la basket
CASE_LINK = "role=link[name='{jpv}'] >> visible=true"                # PLACEHOLDER: link del JPV en la grilla
POST_SEARCH_PAUSE_MS = 2000    # espera después de buscar, antes de clickear el JPV

# --- Caso (ventana principal, dentro de iframes PegaGadget*) ---
REASSIGN_BUTTON = "role=button[name='Reassign'] >> visible=true"    # PLACEHOLDER: confirmar si es botón o menuitem
REASSIGN_TIMEOUT_MS = 15000    # cuánto esperar a que aparezca Reassign en el caso
CANCEL_REASSIGN = "role=button[name='Cancel'] >> visible=true"      # PLACEHOLDER: cierra el diálogo en DRY_RUN

# --- Ejecución ---
DRY_RUN = True                 # True: abre Reassign y cancela sin confirmar. False: reasigna de verdad.
BROWSER_CHANNEL = "msedge"     # "chrome" o "msedge": navegador ya instalado; None: Chromium de Playwright
HEADLESS = False               # False: ves la automatización corriendo
PAUSE_BETWEEN_ROWS_MS = 1500   # pausa entre una fila y la siguiente
SALIDA_CSV = "resultados.csv"

# Colores para la terminal: verde = OK, amarillo = simulado u omitido, rojo = error
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RESET = "\033[0m"
if sys.platform == "win32":
    os.system("")  # activa los colores ANSI en la terminal de Windows

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)-7s %(message)s")
log = logging.getLogger("reasignar_pega")


def to_text(value) -> str:
    """Convierte una celda de Excel a texto. Los enteros no quedan con '.0'."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def read_rows(path: Path):
    """Devuelve [(nro_fila, lob, jpv)] para cada fila del XLSX que tenga algún dato."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[XLSX_SHEET] if XLSX_SHEET else wb.active
        idx_lob = column_index_from_string(COL_LOB) - 1
        idx_jpv = column_index_from_string(COL_JPV) - 1
        first_row = 2 if XLSX_HAS_HEADER else 1
        rows = []
        for nro, row in enumerate(ws.iter_rows(min_row=first_row, values_only=True), start=first_row):
            lob = to_text(row[idx_lob]) if idx_lob < len(row) else ""
            jpv = to_text(row[idx_jpv]) if idx_jpv < len(row) else ""
            if lob or jpv:
                rows.append((nro, lob, jpv))
        return rows
    finally:
        wb.close()


def find_basket(context: BrowserContext, main_page: Page) -> Page:
    """La basket es el popup abierto desde la ventana principal de Pega."""
    popups = [p for p in context.pages if p is not main_page and not p.is_closed()]
    if not popups:
        raise RuntimeError("No encontré la basket: abrila como popup desde Pega y volvé a intentar.")
    return popups[-1]


def buscar_frame(page: Page, locator: str, timeout_ms: int):
    """Busca el iframe de Pega que contiene el locator. Pega carga el caso dentro de PegaGadget*."""
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        for frame in page.frames:
            try:
                if frame.locator(locator).count() > 0:
                    return frame
            except Exception:
                continue
        if time.monotonic() >= deadline:
            return None
        page.wait_for_timeout(500)


def completar_reasignacion(main_page: Page, destino: str) -> None:
    """PENDIENTE: el diálogo de Reassign todavía no está mapeado.

    Falta: el campo donde se elige la basket destino (dropdown, búsqueda o lista)
    y el botón de confirmación. Hasta completarlo, usar DRY_RUN = True.
    """
    raise NotImplementedError("Falta completar el diálogo de Reassign (completar_reasignacion)")


def cancelar_reasignacion(main_page: Page) -> None:
    """Cierra el diálogo de Reassign sin confirmar."""
    frame = buscar_frame(main_page, CANCEL_REASSIGN, 3000)
    if frame is not None:
        frame.locator(CANCEL_REASSIGN).click()
    else:
        main_page.keyboard.press("Escape")


def reasignar_fila(main_page: Page, basket: Page, jpv: str, destino: str, dry_run: bool):
    """Procesa una fila. Devuelve (estado, detalle)."""
    # 1. Buscar el JPV en la basket
    basket.bring_to_front()
    search = basket.locator(BASKET_SEARCH_INPUT)
    search.fill(jpv)
    search.press("Enter")
    basket.wait_for_timeout(POST_SEARCH_PAUSE_MS)

    # 2. Abrir el caso: se carga en la ventana principal
    basket.locator(CASE_LINK.format(jpv=jpv)).click()
    main_page.bring_to_front()

    # 3. Reassign dentro del iframe del caso
    frame = buscar_frame(main_page, REASSIGN_BUTTON, REASSIGN_TIMEOUT_MS)
    if frame is None:
        raise RuntimeError("No apareció el botón Reassign en el caso")
    frame.locator(REASSIGN_BUTTON).click()

    # 4. Basket destino y confirmación
    if dry_run:
        cancelar_reasignacion(main_page)
        return "SIMULADO", f"Destino {destino}"
    completar_reasignacion(main_page, destino)
    return "OK", f"Reasignado a {destino}"


def imprimir(fila, lob, jpv, estado, detalle):
    color = {"OK": GREEN, "ERROR": RED}.get(estado, YELLOW)
    print(f"{color}{estado} - fila {fila}: JPV {jpv} (LOB {lob}) {detalle}{RESET}")


def main():
    if len(sys.argv) != 2:
        print("Uso: python ui_automation.py datos.xlsx")
        sys.exit(1)

    rows = read_rows(Path(sys.argv[1]))
    log.info("Filas a procesar: %d (DRY_RUN=%s)", len(rows), DRY_RUN)
    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(channel=BROWSER_CHANNEL, headless=HEADLESS)
        context = browser.new_context()
        main_page = context.new_page()
        main_page.goto(APP_URL)

        input("Iniciá sesión en Pega, abrí la basket (popup) y después presioná Enter acá... ")
        basket = find_basket(context, main_page)

        for idx, (nro, lob, jpv) in enumerate(rows):
            if basket.is_closed():
                log.error("La basket se cerró. Se marcan las filas restantes como NO PROCESADA.")
                for nro_r, lob_r, jpv_r in rows[idx:]:
                    resultados.append((nro_r, lob_r, jpv_r, "", "NO PROCESADA", "La basket se cerró"))
                break

            destino = MAPA_LOB.get(lob)
            if not jpv:
                resultados.append((nro, lob, jpv, "", "OMITIDA", "JPV vacío"))
                imprimir(nro, lob, jpv, "OMITIDA", "JPV vacío")
                continue
            if destino is None:
                resultados.append((nro, lob, jpv, "", "OMITIDA", f"LOB sin basket: '{lob}'"))
                imprimir(nro, lob, jpv, "OMITIDA", f"LOB sin basket: '{lob}'")
                continue

            try:
                estado, detalle = reasignar_fila(main_page, basket, jpv, destino, DRY_RUN)
            except Exception as e:
                estado, detalle = "ERROR", str(e).splitlines()[0][:200]
                try:
                    main_page.screenshot(path=f"error_fila_{nro}.png")
                except Exception:
                    pass

            resultados.append((nro, lob, jpv, destino, estado, detalle))
            imprimir(nro, lob, jpv, estado, detalle)
            main_page.wait_for_timeout(PAUSE_BETWEEN_ROWS_MS)

        browser.close()

    with open(SALIDA_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["fila_excel", "LOB", "JPV", "basket_destino", "estado", "detalle"])
        w.writerows(resultados)

    total = {}
    for r in resultados:
        total[r[4]] = total.get(r[4], 0) + 1
    log.info("Resumen: %s. Detalle en %s", total, SALIDA_CSV)


if __name__ == "__main__":
    main()
