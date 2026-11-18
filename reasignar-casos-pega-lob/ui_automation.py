#!/usr/bin/env python3
"""
Reasigna casos de Pega a la work queue que corresponde a su LOB, a partir de un XLSX (una pasada por fila).

Por cada fila:
  1. Hace click en el JPV (columna C, "ID") en la basket (popup). El caso se abre en la ventana principal.
  2. Pestaña Work: si el case owner no es TSFX1213, la fila se saltea sin tocar el caso.
  3. Actions > Reassign.
  4. Elige la work queue en el dropdown "Work Queue" según la columna A y, en el caso especial, la columna B
     (ver destino_para).
  5. Submit y Close. En DRY_RUN, en cambio, Close y Discard (no se guarda nada).

Antes de la primera fila, vos hacés el login en la ventana principal, abrís la basket
desde "View Queue for" (popup) y después volvés a la terminal y presionás Enter.

Uso:
    python ui_automation.py datos.xlsx

Por defecto corre con DRY_RUN = True. Genera resultados.csv y una captura (error_fila_N.png)
por cada fila que falle.
"""

import csv
import logging
import os
import re
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
COL_B = "B"                    # columna B: se usa solo en el caso especial
COL_JPV = "C"                  # columna "ID": JPV (ID del caso)
MAPA_LOB = {                   # LOB (columna A) -> value de la work queue en el dropdown
    "Corp&FI": "WDPCTSTWRKFX12",
    "Retail": "WDPCTSTWRKFX13",
}
# Caso especial: se cumplen las dos condiciones a la vez (A y B) -> work queue destino.
# EN PAUSA: con False, las filas "To check" se omiten. Cambiar a True para reactivarlo.
CASO_ESPECIAL_ACTIVO = False
CASO_ESPECIAL = {
    "lob": "To check",
    "col_b": "FX Retail Inv",
    "destino": "WDPCTSTWRKFX13",  # TSFX13
}

# --- Basket (popup) ---
JPV_LINK_NTH = 2               # índice del link del JPV en la basket, según codegen
POST_OPEN_CASE_PAUSE_MS = 2000 # espera a que el caso cargue en la ventana principal

# --- Caso (ventana principal, dentro de iframes PegaGadget*) ---
OWNER_EXPECTED = "TSFX1213"                                       # el case owner tiene que ser exactamente este valor
WORK_TAB = 'text="Work" >> visible=true'                          # pestaña Work del caso
WORK_PANEL = "role=tabpanel[name='Work'] >> visible=true"         # contenido de Work: ahí figura el case owner
ACTIONS_MENU = "role=button[name='Actions'] >> visible=true"
REASSIGN_OPTION = "role=menuitem[name='Reassign'] >> visible=true"
WORK_QUEUE_SELECTOR = 'internal:label="Work Queue"i'              # dropdown de la work queue destino (igual que get_by_label del codegen)
SUBMIT_BUTTON = "role=button[name='Submit'] >> visible=true"
CLOSE_BUTTON = "role=button[name='Close'] >> visible=true"
DISCARD_BUTTON = "role=button[name='Discard'] >> visible=true"    # confirma el descarte en DRY_RUN
PEGA_TIMEOUT_MS = 15000        # cuánto buscar un elemento dentro de los iframes de Pega

# --- Login ---
AUTO_SSO_CLICKS = True         # True: el script hace los clicks de SSO. Usuario, contraseña y MFA los completás vos.
SSO_LINK_NAME = "Login with SSO"
SSO_PROVIDER_NAME = "JPMorgan Chase (Default)"

# --- Abrir la basket ---
AUTO_ABRIR_BASKET = True       # True: el script abre la basket (My teams > View Queue for). False: la abrís a mano.
BASKET_MENU_ITEM_NAME = "My teams"                    # menuitem de la ventana principal (texto, sin el ícono)
BASKET_QUEUE_LABEL = "View Queue for"                 # dropdown que abre la basket como popup
BASKET_QUEUE_VALUE = "WDPCTSXWRKFX1213"               # basket a abrir (confirmado)

# --- Ejecución ---
DRY_RUN = True                 # True: hace todo hasta Close y Discard, sin guardar. False: reasigna de verdad.
BROWSER_CHANNEL = "msedge"     # "chrome" o "msedge": navegador ya instalado; None: Chromium de Playwright
HEADLESS = False               # False: ves la automatización corriendo
MAXIMIZAR_VENTANA = True       # True: abre el navegador maximizado, para que Pega muestre todos los dropdowns
STEP_PAUSE_MS = 500            # pausa breve entre pasos
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


class SkipRow(Exception):
    """La fila no se procesa a propósito (no es un error)."""


def to_text(value) -> str:
    """Convierte una celda de Excel a texto. Los enteros no quedan con '.0'."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def contiene_valor(texto: str, valor: str) -> bool:
    """True si valor aparece como palabra completa: FX1213 no es FX12."""
    return re.search(rf"(?<![A-Za-z0-9]){re.escape(valor)}(?![A-Za-z0-9])", texto) is not None


def login_sso(page: Page) -> None:
    """Clicks de SSO. Usuario, contraseña y MFA los completa la persona."""
    page.get_by_role("link", name=SSO_LINK_NAME, exact=True).click()
    page.get_by_role("button", name=SSO_PROVIDER_NAME).click()
    page.wait_for_timeout(STEP_PAUSE_MS)


def diagnosticar_dropdowns(page: Page) -> str:
    """Lista los dropdowns (select) de cada frame con su label, para ver cómo se llaman en pantalla."""
    partes = []
    for frame in page.frames:
        try:
            labels = frame.evaluate(
                """() => [...document.querySelectorAll('select')].map(s => {
                    const l = (s.labels && s.labels[0]) ? s.labels[0].textContent : (s.getAttribute('aria-label') || s.name || s.id);
                    return (l || '').trim().slice(0, 60);
                })"""
            )
        except Exception:
            continue
        if labels:
            partes.append(f"  frame '{frame.name or 'principal'}': {labels}")
    return "\n".join(partes) or "  ningún dropdown (select) en ningún frame"


def abrir_basket(main_page: Page) -> Page:
    """My teams > View Queue for. La basket se abre como popup y se devuelve esa página."""
    main_page.get_by_role("menuitem", name=re.compile(BASKET_MENU_ITEM_NAME)).click()
    main_page.wait_for_timeout(STEP_PAUSE_MS)
    selector = f'internal:label="{BASKET_QUEUE_LABEL}"i'
    with main_page.expect_popup() as popup_info:
        frame = buscar_frame(main_page, selector, PEGA_TIMEOUT_MS)
        if frame is None:
            main_page.screenshot(path="error_abrir_basket.png")
            raise RuntimeError(
                f"no apareció el dropdown '{BASKET_QUEUE_LABEL}'. Dropdowns encontrados:\n"
                + diagnosticar_dropdowns(main_page)
                + "\nCaptura: error_abrir_basket.png"
            )
        frame.locator(selector).select_option(value=BASKET_QUEUE_VALUE)
    basket = popup_info.value
    basket.wait_for_load_state()
    return basket


def normalizar(texto: str) -> str:
    """Compara textos sin distinguir mayúsculas ni espacios de más."""
    return " ".join(texto.split()).casefold()


def destino_para(lob: str, col_b: str):
    """Work queue destino de la fila, o None si no corresponde.

    El caso especial exige A y B a la vez. Si no se cumple, rige el mapeo de la columna A.
    """
    if (
        CASO_ESPECIAL_ACTIVO
        and normalizar(lob) == normalizar(CASO_ESPECIAL["lob"])
        and normalizar(col_b) == normalizar(CASO_ESPECIAL["col_b"])
    ):
        return CASO_ESPECIAL["destino"]
    for lob_mapeada, destino in MAPA_LOB.items():
        if normalizar(lob) == normalizar(lob_mapeada):
            return destino
    return None


def read_rows(path: Path):
    """Devuelve [(nro_fila, lob, col_b, jpv)] para cada fila del XLSX que tenga algún dato."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[XLSX_SHEET] if XLSX_SHEET else wb.active
        idx_lob = column_index_from_string(COL_LOB) - 1
        idx_b = column_index_from_string(COL_B) - 1
        idx_jpv = column_index_from_string(COL_JPV) - 1
        first_row = 2 if XLSX_HAS_HEADER else 1
        rows = []
        for nro, row in enumerate(ws.iter_rows(min_row=first_row, values_only=True), start=first_row):
            lob = to_text(row[idx_lob]) if idx_lob < len(row) else ""
            col_b = to_text(row[idx_b]) if idx_b < len(row) else ""
            jpv = to_text(row[idx_jpv]) if idx_jpv < len(row) else ""
            if lob or col_b or jpv:
                rows.append((nro, lob, col_b, jpv))
        return rows
    finally:
        wb.close()


def find_basket(context: BrowserContext, main_page: Page) -> Page:
    """La basket es el popup abierto desde la ventana principal de Pega."""
    popups = [p for p in context.pages if p is not main_page and not p.is_closed()]
    if not popups:
        raise RuntimeError("No encontré la basket: abrila como popup desde View Queue for y volvé a intentar.")
    return popups[-1]


def buscar_frame(page: Page, selector: str, timeout_ms: int):
    """Busca el iframe de Pega que contiene el selector. Pega cambia el número del iframe, así que no se fija."""
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        for frame in page.frames:
            try:
                if frame.locator(selector).count() > 0:
                    return frame
            except Exception:
                continue
        if time.monotonic() >= deadline:
            return None
        page.wait_for_timeout(500)


def clic(page: Page, selector: str, paso: str) -> None:
    """Hace click en el elemento dentro del iframe de Pega que lo contiene."""
    frame = buscar_frame(page, selector, PEGA_TIMEOUT_MS)
    if frame is None:
        raise RuntimeError(f"no apareció {paso}")
    frame.locator(selector).click()
    page.wait_for_timeout(STEP_PAUSE_MS)


def reasignar_fila(main_page: Page, basket: Page, jpv: str, destino: str, dry_run: bool):
    """Procesa una fila. Devuelve (estado, detalle). Lanza SkipRow si la fila no corresponde."""
    # 1. Abrir el caso desde la basket: se carga en la ventana principal
    basket.bring_to_front()
    link_jpv = basket.get_by_text(jpv, exact=True)
    if link_jpv.count() <= JPV_LINK_NTH:
        raise SkipRow("el JPV no aparece en la basket")
    link_jpv.nth(JPV_LINK_NTH).click()
    main_page.bring_to_front()
    main_page.wait_for_timeout(POST_OPEN_CASE_PAUSE_MS)

    # 2. Case owner: debe ser TSFX1213
    # PROVISIONAL: busca el valor como palabra completa dentro del panel Work.
    # Cambiar por igualdad exacta cuando tengamos el locator del campo Case owner.
    clic(main_page, WORK_TAB, "pestaña Work")
    frame = buscar_frame(main_page, WORK_PANEL, PEGA_TIMEOUT_MS)
    if frame is None:
        raise RuntimeError("no apareció el panel Work")
    owner = frame.locator(WORK_PANEL).inner_text()
    if not contiene_valor(owner, OWNER_EXPECTED):
        raise SkipRow(f"el case owner no es {OWNER_EXPECTED}")

    # 3. Actions > Reassign
    clic(main_page, ACTIONS_MENU, "botón Actions")
    clic(main_page, REASSIGN_OPTION, "opción Reassign")

    # 4. Work queue destino según LOB
    frame = buscar_frame(main_page, WORK_QUEUE_SELECTOR, PEGA_TIMEOUT_MS)
    if frame is None:
        raise RuntimeError("no apareció el dropdown Work Queue")
    frame.locator(WORK_QUEUE_SELECTOR).select_option(value=destino)
    main_page.wait_for_timeout(STEP_PAUSE_MS)

    # 5. Guardar y cerrar (o descartar en DRY_RUN)
    if dry_run:
        clic(main_page, CLOSE_BUTTON, "botón Close")
        clic(main_page, DISCARD_BUTTON, "botón Discard")
        return "SIMULADO", f"Destino {destino}"
    clic(main_page, SUBMIT_BUTTON, "botón Submit")
    clic(main_page, CLOSE_BUTTON, "botón Close")
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
    # Se valida antes de abrir el navegador: el caso especial necesita que A y B coincidan.
    resultados = []

    with sync_playwright() as p:
        args = ["--start-maximized"] if MAXIMIZAR_VENTANA else []
        browser = p.chromium.launch(channel=BROWSER_CHANNEL, headless=HEADLESS, args=args)
        # no_viewport: la página usa el tamaño real de la ventana, así que queda maximizada
        context = browser.new_context(no_viewport=MAXIMIZAR_VENTANA)
        main_page = context.new_page()
        main_page.goto(APP_URL)
        if AUTO_SSO_CLICKS:
            login_sso(main_page)

        if AUTO_ABRIR_BASKET:
            input("Completá usuario, contraseña y MFA. Cuando se vea Pega, presioná Enter acá... ")
            basket = abrir_basket(main_page)
        else:
            input("Completá usuario, contraseña y MFA, abrí la basket desde 'View Queue for' (popup) y después presioná Enter acá... ")
            basket = find_basket(context, main_page)

        for idx, (nro, lob, col_b, jpv) in enumerate(rows):
            if basket.is_closed():
                log.error("La basket se cerró. Se marcan las filas restantes como NO PROCESADA.")
                for nro_r, lob_r, col_b_r, jpv_r in rows[idx:]:
                    resultados.append((nro_r, lob_r, jpv_r, "", "NO PROCESADA", "La basket se cerró"))
                break

            destino = destino_para(lob, col_b)
            if not jpv:
                resultados.append((nro, lob, jpv, "", "OMITIDA", "JPV vacío"))
                imprimir(nro, lob, jpv, "OMITIDA", "JPV vacío")
                continue
            if destino is None:
                detalle = f"sin work queue: A='{lob}', B='{col_b}'"
                resultados.append((nro, lob, jpv, "", "OMITIDA", detalle))
                imprimir(nro, lob, jpv, "OMITIDA", detalle)
                continue

            try:
                estado, detalle = reasignar_fila(main_page, basket, jpv, destino, DRY_RUN)
            except SkipRow as e:
                estado, detalle = "OMITIDA", str(e)
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
        w.writerow(["fila_excel", "LOB", "JPV", "work_queue_destino", "estado", "detalle"])
        w.writerows(resultados)

    total = {}
    for r in resultados:
        total[r[4]] = total.get(r[4], 0) + 1
    log.info("Resumen: %s. Detalle en %s", total, SALIDA_CSV)


if __name__ == "__main__":
    main()
