#!/usr/bin/env python3
"""
Asigna casos a cada empleado en Pega a partir de un XLSX (una pasada por fila).

Por cada fila:
  1. Busca el JPV (columna B) en la search bar y confirma la búsqueda.
  2. Si el tipo de caso (columna D) es Miscellaneous: Actions > Change Case Type,
     elige Amendment en el dropdown y hace Submit.
  3. Abre Actions y elige Reassign.
  4. Elige el assignee (columna A) en el dropdown "operator".
  5. Hace click en Submit y Close. No hace falta Save: al reasignar no hay cambios que guardar.

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
import os
import re
import sys
import time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from playwright.sync_api import FrameLocator, Page, sync_playwright

# =====================================================================
# CONFIGURACIÓN: completar con los selectores reales de la app.
# Tip: `playwright codegen <URL>` graba los clicks y devuelve selectores.
# Los valores de ejemplo son placeholders.
# =====================================================================
APP_URL = "https://TU-APP-INTERNA"
try:
    # URL real, en config_local.py (ver config_local.example.py). No se sube al repo.
    from config_local import APP_URL
except ImportError:
    pass

# --- Búsqueda por JPV ---
SEARCH_INPUT = "role=textbox[name='Search']"       # search bar donde va el JPV
SEARCH_BUTTON = "role=button[name='Search']"        # lupa (se usa si SEARCH_WITH_ENTER es False)
SEARCH_WITH_ENTER = True                            # True: Enter confirma la búsqueda; False: click en la lupa

# --- Acciones sobre el caso (dentro del iframe de Pega) ---
PEGA_FRAME_PREFIX = "PegaGadget"                    # Pega cambia el número del iframe (PegaGadget1Ifr, PegaGadget2Ifr...)
PEGA_FRAME_TIMEOUT_MS = 15000                       # cuánto buscar el iframe con el botón Actions
ACTIONS_MENU = "role=button[name='Actions'] >> visible=true"        # PegaGadget Actions
REASSIGN_OPTION = "role=menuitem[name='Reassign'] >> visible=true"  # PegaGadget Reassign
# Step By Step: puede ser menuitem o botón, y puede tener un nombre de empleado a continuación
STEP_BY_STEP_OPTIONS = (
    "role=menuitem[name=/^Step By Step/i] >> visible=true",
    "role=button[name=/^Step By Step/i] >> visible=true",
)
REOPEN_BUTTON = "role=button[name=/^Reopen/i] >> visible=true"  # si aparece, el caso no se reasigna: se saltea
MESSAGE_ATTACHED_OPTIONS = (  # último recurso si no aparece ninguna opción anterior
    "role=menuitem[name=/^MessageAttached/i] >> visible=true",
    "role=button[name=/^MessageAttached/i] >> visible=true",
)
PERFORM_RESEARCH_OPTION = "role=menuitem[name=/^Perform research/i] >> visible=true"  # opción que se usa si tampoco aparece Step By Step
MENU_CHECK_MS = 2000                                # cuánto esperar a que aparezca una opción del menú Actions
OPERATOR_LABEL = "operator"                         # label del dropdown "operator": ahí se elige el assignee
SUBMIT_BUTTON = "role=button[name='Submit'] >> visible=true"        # PegaGadget Submit
CLOSE_BUTTON = "role=button[name='Close'] >> visible=true"          # PegaGadget Close
WORK_TAB = 'text="Work" >> visible=true'                            # pestaña Work del caso (texto exacto, como en codegen)
WORK_PANEL = "role=tabpanel[name='Work'] >> visible=true"           # contenido de la pestaña Work: ahí figura el asignado
OWNERS_ALLOWED = ("FX12", "AutoCV")                 # el reassign solo se hace si el campo Work dice alguno de estos
BOT_ID = "B089707"                                 # si el asignado actual contiene este ID, la fila se saltea
CHANGE_CASE_TYPE_OPTION = "role=menuitem[name='Change case type'] >> visible=true"  # opción dentro de Actions
NEW_CASE_TYPE_LABEL = "Select New Case Type"        # label del dropdown donde se elige el nuevo tipo de caso

XLSX_SHEET = "Cases Distribution"  # nombre exacto de la hoja. Si el XLSX no la tiene, el script se detiene antes de abrir el navegador
XLSX_HAS_HEADER = True     # True: la primera fila es encabezado y se saltea
COL_ASSIGNEE = "A"         # columna con el assignee
COL_JPV = "B"              # columna con el identificador JPV
COL_TEAM = "C"             # columna con el equipo
TEAMS_ALLOWED = {"AUTOCV", "FX ChaseTr", "iFX"}  # solo se procesan las filas de estos equipos; el resto se omite
COL_CASE_TYPE = "D"        # columna con el tipo de caso
CASE_TYPE_TRIGGER = "Miscellaneous"  # si la fila tiene este tipo, antes de reasignar se cambia el tipo de caso
CASE_TYPES_SOLO_AUTOCV = {"Miscellaneous", "OFAC"}  # estos tipos solo se procesan si el equipo es AUTOCV; si no, se omiten
TEAM_AUTOCV = "AUTOCV"
NEW_CASE_TYPE = "JPM-Work-Shared-WDPC-Exp-AmendBen"  # VALUE de la opción en el dropdown (no el texto visible)

# Nombres del XLSX que no coinciden con el dropdown de la app: {nombre_en_xlsx: nombre_en_app}
NAME_ALIASES = {
    "Iara Mañani": "Iara Belen Manani",
    "Luciana Scarone": "Luciana Veronica Scarone",
}

BROWSER_CHANNEL = "msedge"  # "chrome" o "msedge": usa el navegador ya instalado; None: Chromium de Playwright
HEADLESS = False        # False: ves la automatización corriendo
AUTO_SSO_CLICKS = True  # True: el script hace click en "Login with SSO" y en "JPMorgan Chase (Default)"
LOGIN_TIMEOUT_MS = 120000  # cuánto esperar a que termines usuario, contraseña y MFA (la barra de búsqueda aparece al terminar)
STEP_PAUSE_MS = 300     # pausa breve entre pasos
PAUSE_BETWEEN_ROWS_MS = 1500  # pausa mínima entre una fila y la siguiente
POST_SEARCH_PAUSE_MS = 2000   # espera después de buscar el JPV, antes de tocar el caso (Work, Actions)
RECOVER_TIMEOUT_MS = 15000  # cuánto esperar la search bar al recuperarse de un error

# Colores para la terminal: verde = OK, rojo = error
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RESET = "\033[0m"
if sys.platform == "win32":
    os.system("")  # activa los colores ANSI en la terminal de Windows

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
    """Devuelve [(nro_fila, assignee, jpv, equipo, tipo_caso)] para cada fila válida del XLSX."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if XLSX_SHEET not in wb.sheetnames:
            sys.exit(f"El XLSX no tiene la hoja '{XLSX_SHEET}'. Hojas encontradas: {', '.join(wb.sheetnames)}")
        ws = wb[XLSX_SHEET]
        idx_assignee = column_index_from_string(COL_ASSIGNEE) - 1
        idx_jpv = column_index_from_string(COL_JPV) - 1
        idx_team = column_index_from_string(COL_TEAM) - 1
        idx_case_type = column_index_from_string(COL_CASE_TYPE) - 1
        first_row = 2 if XLSX_HAS_HEADER else 1

        rows = []
        for row_no, cells in enumerate(
            ws.iter_rows(min_row=first_row, values_only=True), start=first_row
        ):
            assignee = to_text(cells[idx_assignee]) if idx_assignee < len(cells) else ""
            assignee = NAME_ALIASES.get(assignee, assignee)
            jpv = to_text(cells[idx_jpv]) if idx_jpv < len(cells) else ""
            team = to_text(cells[idx_team]) if idx_team < len(cells) else ""
            case_type = to_text(cells[idx_case_type]) if idx_case_type < len(cells) else ""
            if not assignee or not jpv:
                if assignee or jpv:  # fila incompleta: se avisa, no se procesa
                    log.warning("Fila %d omitida (falta assignee o JPV)", row_no)
                continue
            rows.append((row_no, assignee, jpv, team, case_type))
        return rows
    finally:
        wb.close()


def login_sso(page: Page) -> None:
    """Clicks de SSO. Usuario, contraseña y MFA los completa la persona."""
    page.get_by_role("link", name="Login with SSO", exact=True).click()
    page.get_by_role("button", name="JPMorgan Chase (Default)").click()
    page.wait_for_timeout(STEP_PAUSE_MS)


class StepError(Exception):
    """Error en un paso de una fila. El mensaje incluye el nombre del paso."""


class SkipRow(Exception):
    """La fila no se procesa a propósito (no es un error)."""


def contiene_valor(texto: str, valor: str) -> bool:
    """True si valor aparece como palabra completa: FX1213 no es FX12, B0897071 no es B089707."""
    return re.search(rf"(?<![A-Za-z0-9]){re.escape(valor)}(?![A-Za-z0-9])", texto) is not None


def menu_option_visible(pega: FrameLocator, selector: str) -> bool:
    """True si la opción aparece en el menú Actions dentro de MENU_CHECK_MS."""
    try:
        pega.locator(selector).wait_for(state="visible", timeout=MENU_CHECK_MS)
        return True
    except Exception:  # noqa: BLE001
        return False


def find_frame_with(page: Page, selector: str) -> FrameLocator:
    """Devuelve el iframe de Pega que contiene el elemento indicado.

    Pega cambia el número del iframe (PegaGadget1Ifr, PegaGadget2Ifr...)
    después de un error o una recarga, así que no se fija el nombre.
    """
    for _ in range(max(PEGA_FRAME_TIMEOUT_MS // STEP_PAUSE_MS, 1)):
        for frame in page.frames:
            if frame.name.startswith(PEGA_FRAME_PREFIX) and frame.locator(selector).count() > 0:
                return page.frame_locator(f"iframe[name='{frame.name}']")
        page.wait_for_timeout(STEP_PAUSE_MS)
    raise TimeoutError(f"no apareció {selector} en ningún iframe de Pega")


def find_pega_frame(page: Page) -> FrameLocator:
    """Iframe de Pega que contiene el botón Actions."""
    return find_frame_with(page, ACTIONS_MENU)


def valor_opcion(select, nombre: str) -> str:
    """Devuelve el value de la opción del dropdown cuyo texto es el nombre, sin distinguir mayúsculas ni espacios de más.

    Reintenta hasta PEGA_FRAME_TIMEOUT_MS: las opciones pueden cargarse después de abrir el dropdown.
    """
    buscado = " ".join(nombre.split()).casefold()
    limite = time.monotonic() + PEGA_FRAME_TIMEOUT_MS / 1000
    while True:
        opciones = select.locator("option").all()
        for opcion in opciones:
            texto = " ".join((opcion.text_content() or "").split()).casefold()
            if texto == buscado:
                return opcion.get_attribute("value") or (opcion.text_content() or "").strip()
        if time.monotonic() >= limite:
            leidas = [" ".join((o.text_content() or "").split()) for o in opciones[:10]]
            raise ValueError(
                f"no encontré '{nombre}' en el dropdown (opciones leídas: {len(opciones)}; primeras: {leidas})"
            )
        time.sleep(STEP_PAUSE_MS / 1000)


def process_row(page: Page, assignee: str, jpv: str, case_type: str) -> None:
    paso = "buscar JPV"
    try:
        # 1) Buscar el JPV
        search = page.locator(SEARCH_INPUT)
        search.click()
        search.fill(jpv)
        escrito = search.input_value()
        if escrito != jpv:
            raise ValueError(f"la search bar tiene '{escrito}' en vez de '{jpv}'")
        if SEARCH_WITH_ENTER:
            search.press("Enter")
        else:
            page.locator(SEARCH_BUTTON).click()
        page.wait_for_timeout(POST_SEARCH_PAUSE_MS)

        # Pestaña Work y verificación del asignado: si es el bot, la fila se saltea
        paso = "pestaña Work"
        find_frame_with(page, WORK_TAB).locator(WORK_TAB).click()
        page.wait_for_timeout(STEP_PAUSE_MS)

        paso = "iframe de Pega (botón Actions)"
        pega = find_pega_frame(page)
        paso = "leer asignado actual"
        owner = find_frame_with(page, WORK_PANEL).locator(WORK_PANEL).inner_text(timeout=PEGA_FRAME_TIMEOUT_MS)
        if not owner.strip():
            raise ValueError("no se pudo leer el asignado actual")
        if contiene_valor(owner, BOT_ID):
            raise SkipRow(f"asignado al bot {BOT_ID}")
        if not any(contiene_valor(owner, o) for o in OWNERS_ALLOWED):
            raise SkipRow(f"no se reasigna: el campo Work no dice {' ni '.join(OWNERS_ALLOWED)}")
        if menu_option_visible(pega, REOPEN_BUTTON):
            raise SkipRow("tiene el botón Reopen: no se reasigna")


        # 2) Actions. Solo para Miscellaneous y si aparece Change case type:
        #    Change case type > Amendment > Submit, y vuelve a abrir Actions.
        paso = "Actions"
        pega.locator(ACTIONS_MENU).click()
        if case_type == CASE_TYPE_TRIGGER and menu_option_visible(pega, CHANGE_CASE_TYPE_OPTION):
            paso = "Change case type"
            pega.locator(CHANGE_CASE_TYPE_OPTION).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = "Select New Case Type"
            pega.get_by_label(NEW_CASE_TYPE_LABEL).select_option(value=NEW_CASE_TYPE)
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = "Submit (cambio de tipo)"
            pega.locator(SUBMIT_BUTTON).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = "Actions (tras cambio de tipo)"
            pega.locator(ACTIONS_MENU).click()
        elif case_type == CASE_TYPE_TRIGGER:
            log.info("Tipo Miscellaneous sin Change case type en el menú: sigue con Reassign")

        # 3) Reassign (si no aparece: Step By Step, y después Reassign)
        if not menu_option_visible(pega, REASSIGN_OPTION):
            step_by_step = next((sel for sel in STEP_BY_STEP_OPTIONS if menu_option_visible(pega, sel)), None)
            if step_by_step:
                paso = "Step By Step"
                pega.locator(step_by_step).click()
            elif menu_option_visible(pega, PERFORM_RESEARCH_OPTION):
                paso = "Perform research"
                pega.locator(PERFORM_RESEARCH_OPTION).click()
            else:
                paso = "MessageAttached"
                message = next((sel for sel in MESSAGE_ATTACHED_OPTIONS if menu_option_visible(pega, sel)), None)
                if message is None:
                    raise TimeoutError("no apareció Reassign, Step By Step, Perform research ni MessageAttached")
                pega.locator(message).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
        paso = "Reassign"
        pega.locator(REASSIGN_OPTION).click()
        page.wait_for_timeout(STEP_PAUSE_MS)

        # 4) Elegir el assignee en el dropdown "operator" (sin importar mayúsculas ni espacios)
        paso = "operator (assignee)"
        operator = pega.get_by_label(OPERATOR_LABEL)
        operator.select_option(value=valor_opcion(operator, assignee))
        page.wait_for_timeout(STEP_PAUSE_MS)

        # 5) Submit y Close (sin Save: la reasignación no deja nada por guardar)
        paso = "Submit"
        pega.locator(SUBMIT_BUTTON).click()
        page.wait_for_timeout(STEP_PAUSE_MS)
        paso = "Close"
        pega.locator(CLOSE_BUTTON).click()
        page.wait_for_timeout(STEP_PAUSE_MS)
    except SkipRow:
        raise
    except Exception as exc:
        raise StepError(f"paso '{paso}': {exc}") from exc


def recover(page: Page) -> bool:
    """Después de un error, vuelve a la search bar. Devuelve False si no lo logra."""
    try:
        page.keyboard.press("Escape")
        page.keyboard.press("Escape")
        page.reload()
        page.locator(SEARCH_INPUT).wait_for(state="visible", timeout=RECOVER_TIMEOUT_MS)
        return True
    except Exception:  # noqa: BLE001
        return False


def guardar_resultados(results: list) -> None:
    """Escribe resultados.csv (sin las omitidas) y muestra la tabla de errores y el resumen.

    Se llama también si la corrida se corta a mitad.
    """
    out = Path("resultados.csv")
    filas = [r for r in results if r[3] != "OMITIDA"]
    with out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["fila", "assignee", "jpv", "estado", "detalle"])
        writer.writerows(filas)

    ok = [r for r in filas if r[3] == "OK"]
    errores = [r for r in filas if r[3] == "ERROR"]
    sin_procesar = [r for r in filas if r[3] == "NO PROCESADA"]
    if errores:
        # El ancho sale del dato más largo, con 2 espacios de separación, para que el JPV no pegue con el assignee
        ancho_jpv = max(len("JPV"), *(len(r[2]) for r in errores)) + 2
        ancho_asig = max(len("Asignado a"), *(len(r[1]) for r in errores)) + 2
        print(f"\n{RED}{'JPV':<{ancho_jpv}}{'Asignado a':<{ancho_asig}}Motivo{RESET}")
        for r in errores:
            print(f"{RED}{r[2]:<{ancho_jpv}}{r[1]:<{ancho_asig}}{r[4]}{RESET}")
        print()
    log.info("Resultados: %d OK, %d con error, %d sin procesar. Detalle en %s",
             len(ok), len(errores), len(sin_procesar), out.resolve())


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
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(channel=BROWSER_CHANNEL, headless=HEADLESS, args=["--start-maximized"])
            page = browser.new_page(no_viewport=True)
            page.goto(APP_URL)

            if AUTO_SSO_CLICKS:
                login_sso(page)

            print("Completá usuario, contraseña y MFA en el navegador. El script arranca solo cuando aparezca la búsqueda.")
            page.locator(SEARCH_INPUT).wait_for(state="visible", timeout=LOGIN_TIMEOUT_MS)

            detenido = False
            for row_no, assignee, jpv, team, case_type in rows:
                if detenido:
                    results.append([row_no, assignee, jpv, "NO PROCESADA", "corrida detenida por un error anterior"])
                    continue
                if team not in TEAMS_ALLOWED:
                    motivo = f"equipo '{team}' fuera de TEAMS_ALLOWED"
                    log.warning(f"{YELLOW}JPV %s: %s. Se saltea la fila.{RESET}", jpv, motivo)
                    results.append([row_no, assignee, jpv, "OMITIDA", motivo])
                    continue
                if case_type.casefold() in {t.casefold() for t in CASE_TYPES_SOLO_AUTOCV} and team != TEAM_AUTOCV:
                    motivo = f"tipo '{case_type}' solo se procesa para el equipo {TEAM_AUTOCV}"
                    log.warning(f"{YELLOW}JPV %s: %s. Se saltea la fila.{RESET}", jpv, motivo)
                    results.append([row_no, assignee, jpv, "OMITIDA", motivo])
                    continue
                try:
                    process_row(page, assignee, jpv, case_type)
                    log.info(f"{GREEN}OK - JPV %s asignado a %s{RESET}", jpv, assignee)
                    results.append([row_no, assignee, jpv, "OK", ""])
                except SkipRow as skip:
                    log.warning(f"{YELLOW}JPV %s: %s. Se saltea la fila.{RESET}", jpv, skip)
                    results.append([row_no, assignee, jpv, "OMITIDA", str(skip)])
                except Exception as exc:  # noqa: BLE001
                    text = str(exc).strip()
                    msg = text.splitlines()[0] if text else type(exc).__name__
                    log.error(f"{RED}ERROR - JPV %s no se pudo asignar a %s. Motivo: %s{RESET}", jpv, assignee, msg)
                    try:
                        page.screenshot(path=f"error_fila_{row_no}.png", full_page=True)
                    except Exception:  # la página pudo quedar en mal estado
                        pass
                    results.append([row_no, assignee, jpv, "ERROR", msg])
                    if recover(page):
                        log.info("Volvió a la búsqueda; sigue con la siguiente fila.")
                    else:
                        log.error(f"{RED}No se pudo volver a la búsqueda. Se detiene la corrida.{RESET}")
                        detenido = True
                if not detenido:
                    page.wait_for_timeout(PAUSE_BETWEEN_ROWS_MS)
            browser.close()
    finally:
        guardar_resultados(results)


if __name__ == "__main__":
    main()
