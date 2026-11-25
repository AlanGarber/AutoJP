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

import argparse
import csv
import logging
import os
import re
import subprocess
import sys
import time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from playwright.sync_api import FrameLocator, Page, sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeout

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
OWNER_NO_REASIGNAR = "AutoCV Automation"           # si el campo Work dice esto, la fila se saltea aunque diga AutoCV
BOT_ID = "B089707"                                 # si el asignado actual contiene este ID, la fila se saltea
CHANGE_CASE_TYPE_OPTION = "role=menuitem[name='Change case type'] >> visible=true"  # opción dentro de Actions
# Flujo ReviewPymtPaidCorrectly: Actions > ReviewPymtPaidCorrectly > Actions > Cancel this Assignment > Submit, y después el flujo normal
REVIEW_PYMT_OPTIONS = (
    "role=menuitem[name=/^ReviewPymtPaidCorrectly/i] >> visible=true",
    "role=button[name=/^ReviewPymtPaidCorrectly/i] >> visible=true",
)
CANCEL_ASSIGNMENT_OPTIONS = (
    "role=menuitem[name=/^Cancel this Assignment/i] >> visible=true",
    "role=button[name=/^Cancel this Assignment/i] >> visible=true",
)
NEW_CASE_TYPE_LABEL = "Select New Case Type"        # label del dropdown donde se elige el nuevo tipo de caso

XLSX_SHEETS = ("Cases Distribution", "Distribution")  # nombres exactos aceptados, en orden de preferencia. Si el XLSX no tiene ninguno, el script se detiene antes de abrir el navegador
XLSX_HAS_HEADER = True     # True: la primera fila es encabezado y se saltea
COL_ASSIGNEE = "A"         # columna con el assignee
COL_JPV = "B"              # columna con el identificador JPV
COL_TEAM = "C"             # columna con el equipo
TEAMS_ALLOWED = {"AUTOCV", "FX ChaseTr", "iFX", "FX Retail Inv"}  # solo se procesan las filas de estos equipos; el resto se omite
# Owners exigidos por equipo. Si el equipo no está acá, se usa OWNERS_ALLOWED.
# FX Retail Inv (2026-11-25): el campo Work tiene que decir FX13.
OWNERS_POR_EQUIPO = {"FX Retail Inv": ("FX13",)}
COL_CASE_TYPE = "D"        # columna con el tipo de caso
# Archivo failed: si el nombre del archivo contiene esta palabra, el script lo lee como failed.
# Columnas: A = Investigator (assignee), B = JPV, E = Sub LOB. No tiene equipo ni tipo de caso.
FAILED_FILE_MARK = "failed"
COL_SUB_LOB = "E"          # solo en failed: columna con el Sub LOB
FAILED_SUB_LOBS = ("12", "13")  # solo en failed: se procesan las filas cuyo Sub LOB dice alguno de estos (palabra completa). Antes solo "12"
# Owners exigidos en failed por Sub LOB (2026-11-25): el Sub LOB 13 tiene que tener FX13 en el campo Work. El resto usa OWNERS_ALLOWED.
OWNERS_POR_SUB_LOB = {"13": ("FX13",)}
CASE_TYPE_TRIGGER = "Miscellaneous"  # si la fila tiene este tipo, antes de reasignar se cambia el tipo de caso
CASE_TYPES_SOLO_AUTOCV = {"Miscellaneous", "OFAC"}  # estos tipos solo se procesan si el equipo es AUTOCV; si no, se omiten
EXCEPCIONES_SOLO_AUTOCV = {("iFX", "Miscellaneous")}  # (equipo, tipo) que sí se procesan aunque el equipo no sea AUTOCV
EQUIPOS_TODOS_LOS_TIPOS = {"FX Retail Inv"}  # estos equipos se reasignan con cualquier tipo de caso (2026-11-25)
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
STEP_PAUSE_MS = 200     # pausa breve entre pasos (antes 300; los clicks ya esperan a que el elemento esté listo)
INSTANCIAS = 2                 # cuántas ventanas de Edge en paralelo (se puede cambiar con --instancias N)
ESPERA_ENTRE_INSTANCIAS_MS = 15000  # pausa entre el arranque de cada ventana (evita que dos logins SSO se pisen)
RESULTADOS_CSV = "resultados.csv"
PAUSE_BETWEEN_ROWS_MS = 1500  # pausa mínima entre una fila y la siguiente
POST_SEARCH_PAUSE_MS = 1000   # espera después de buscar el JPV, antes de tocar el caso (antes 2000; si vuelve a tocar el caso anterior, subirla)
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


def read_rows(path: Path, failed: bool = False):
    """Devuelve [(nro_fila, assignee, jpv, equipo, tipo_caso)] para cada fila válida del XLSX.

    En failed, el cuarto campo es el Sub LOB (columna E) y el tipo de caso queda vacío.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        hoja = next((h for h in XLSX_SHEETS if h in wb.sheetnames), None)
        if hoja is None:
            sys.exit(f"El XLSX no tiene ninguna de las hojas {', '.join(repr(h) for h in XLSX_SHEETS)}. "
                     f"Hojas encontradas: {', '.join(wb.sheetnames)}")
        log.info("Usando la hoja '%s'", hoja)
        ws = wb[hoja]
        idx_assignee = column_index_from_string(COL_ASSIGNEE) - 1
        idx_jpv = column_index_from_string(COL_JPV) - 1
        if failed:
            idx_team = column_index_from_string(COL_SUB_LOB) - 1
            idx_case_type = None
        else:
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
            case_type = ""
            if idx_case_type is not None and idx_case_type < len(cells):
                case_type = to_text(cells[idx_case_type])
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


def contiene_frase(texto: str, frase: str) -> bool:
    """True si la frase aparece como palabra completa, sin importar mayúsculas ni cómo se partan los espacios (salto de línea incluido)."""
    partes = r"\s+".join(re.escape(p) for p in frase.split())
    return re.search(rf"(?<![A-Za-z0-9]){partes}(?![A-Za-z0-9])", texto, re.IGNORECASE) is not None


def solo_autocv(team: str, case_type: str) -> bool:
    """True si la fila tiene un tipo que solo se procesa para AUTOCV y su equipo no tiene excepción. Sin distinguir mayúsculas."""
    equipo, tipo = team.casefold(), case_type.casefold()
    if equipo in {e.casefold() for e in EQUIPOS_TODOS_LOS_TIPOS}:
        return False
    if tipo not in {t.casefold() for t in CASE_TYPES_SOLO_AUTOCV}:
        return False
    if equipo == TEAM_AUTOCV.casefold():
        return False
    return (equipo, tipo) not in {(e.casefold(), t.casefold()) for e, t in EXCEPCIONES_SOLO_AUTOCV}


def menu_option_visible(pega: FrameLocator, selector: str) -> bool:
    """True si la opción aparece en el menú Actions dentro de MENU_CHECK_MS."""
    try:
        pega.locator(selector).wait_for(state="visible", timeout=MENU_CHECK_MS)
        return True
    except Exception:  # noqa: BLE001
        return False


def esperar_alguna_visible(pega: FrameLocator, selectores: list) -> bool:
    """True si alguno de los selectores aparece dentro de MENU_CHECK_MS. Espera una sola vez para todos.

    Los selectores ya tienen '>> visible=true', así que cada uno solo matchea elementos visibles.
    """
    union = pega.locator(selectores[0])
    for sel in selectores[1:]:
        union = union.or_(pega.locator(sel))
    try:
        union.first.wait_for(state="visible", timeout=MENU_CHECK_MS)
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
    # Espacios entre palabras se toleran; mayúsculas, no importan. Playwright espera la opción sin polling manual.
    patron = re.compile(r"^\s*" + r"\s+".join(re.escape(p) for p in nombre.split()) + r"\s*$", re.IGNORECASE)
    opcion = select.locator("option").filter(has_text=patron).first
    try:
        opcion.wait_for(state="attached", timeout=PEGA_FRAME_TIMEOUT_MS)
    except PlaywrightTimeout:
        primeras = select.evaluate("s => Array.from(s.options).map(o => o.textContent.trim()).slice(0, 10)")
        raise ValueError(f"no encontré '{nombre}' en el dropdown (primeras opciones: {primeras})")
    return opcion.get_attribute("value") or opcion.text_content().strip()


class Cronometro:
    """Mide cuánto tarda cada paso de una fila. Cada vez que cambia el paso, cierra el anterior."""

    def __init__(self) -> None:
        self.reiniciar()

    def reiniciar(self) -> None:
        self.paso_actual = None
        self.inicio_paso = time.monotonic()
        self.inicio_fila = self.inicio_paso
        self.tiempos = []

    def marcar(self, nombre: str) -> str:
        ahora = time.monotonic()
        if self.paso_actual is not None:
            self.tiempos.append((self.paso_actual, ahora - self.inicio_paso))
        self.paso_actual = nombre
        self.inicio_paso = ahora
        return nombre

    def resumen(self) -> str:
        ahora = time.monotonic()
        tiempos = list(self.tiempos)
        if self.paso_actual is not None:
            tiempos.append((self.paso_actual, ahora - self.inicio_paso))
        lentos = sorted(tiempos, key=lambda t: t[1], reverse=True)[:4]
        detalle = ", ".join(f"{nombre} {segundos:.1f} s" for nombre, segundos in lentos)
        return f"{ahora - self.inicio_fila:.1f} s en total. Más lentos: {detalle}"


CRONO = Cronometro()


def process_row(page: Page, assignee: str, jpv: str, case_type: str, owners: tuple = OWNERS_ALLOWED) -> None:
    CRONO.reiniciar()
    paso = CRONO.marcar("buscar JPV")
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
        paso = CRONO.marcar("pestaña Work")
        find_frame_with(page, WORK_TAB).locator(WORK_TAB).click()
        page.wait_for_timeout(STEP_PAUSE_MS)

        paso = CRONO.marcar("iframe de Pega (botón Actions)")
        pega = find_pega_frame(page)
        paso = CRONO.marcar("leer asignado actual")
        owner = find_frame_with(page, WORK_PANEL).locator(WORK_PANEL).inner_text(timeout=PEGA_FRAME_TIMEOUT_MS)
        if not owner.strip():
            raise ValueError("no se pudo leer el asignado actual")
        if contiene_valor(owner, BOT_ID):
            raise SkipRow(f"asignado al bot {BOT_ID}")
        if contiene_frase(owner, OWNER_NO_REASIGNAR):
            raise SkipRow(f"no se reasigna. Campo work dice {OWNER_NO_REASIGNAR}")
        if not any(contiene_valor(owner, o) for o in owners):
            raise SkipRow(f"no se reasigna: el campo Work no dice {' ni '.join(owners)}")
        if menu_option_visible(pega, REOPEN_BUTTON):
            raise SkipRow("tiene el botón Reopen: no se reasigna")


        # 2) Actions. Solo para Miscellaneous y si aparece Change case type:
        #    Change case type > Amendment > Submit, y vuelve a abrir Actions.
        paso = CRONO.marcar("Actions")
        pega.locator(ACTIONS_MENU).click()

        # Una sola espera para todo el menú: en cuanto aparece cualquier opción, se mira lo demás sin esperar de nuevo.
        # Si no aparece nada, el menú no se abrió y se corta acá.
        opciones_menu = [*REVIEW_PYMT_OPTIONS, CHANGE_CASE_TYPE_OPTION, REASSIGN_OPTION,
                         *MESSAGE_ATTACHED_OPTIONS, *STEP_BY_STEP_OPTIONS, PERFORM_RESEARCH_OPTION]
        if not esperar_alguna_visible(pega, opciones_menu):
            raise TimeoutError("el menú Actions no mostró ninguna opción")

        # 2a) ReviewPymtPaidCorrectly: Actions > ReviewPymtPaidCorrectly > Actions > Cancel this Assignment > Submit.
        #     Después el Actions queda abierto y sigue el flujo normal.
        if any(pega.locator(s).count() > 0 for s in REVIEW_PYMT_OPTIONS):
            sel_review = next(s for s in REVIEW_PYMT_OPTIONS if pega.locator(s).count() > 0)
            paso = CRONO.marcar("ReviewPymtPaidCorrectly")
            pega.locator(sel_review).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = CRONO.marcar("Actions (antes de Cancel this Assignment)")
            pega.locator(ACTIONS_MENU).click()
            paso = CRONO.marcar("Cancel this Assignment")
            if not esperar_alguna_visible(pega, list(CANCEL_ASSIGNMENT_OPTIONS)):
                raise TimeoutError("no apareció Cancel this Assignment en el menú Actions")
            sel_cancel = next(s for s in CANCEL_ASSIGNMENT_OPTIONS if pega.locator(s).count() > 0)
            pega.locator(sel_cancel).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = CRONO.marcar("Submit (Cancel this Assignment)")
            pega.locator(SUBMIT_BUTTON).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = CRONO.marcar("Actions (tras Cancel this Assignment)")
            pega.locator(ACTIONS_MENU).click()
            page.wait_for_timeout(STEP_PAUSE_MS)

        es_trigger = case_type.casefold() == CASE_TYPE_TRIGGER.casefold()
        if es_trigger and pega.locator(CHANGE_CASE_TYPE_OPTION).count() > 0:
            paso = CRONO.marcar("Change case type")
            pega.locator(CHANGE_CASE_TYPE_OPTION).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = CRONO.marcar("Select New Case Type")
            pega.get_by_label(NEW_CASE_TYPE_LABEL).select_option(value=NEW_CASE_TYPE)
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = CRONO.marcar("Submit (cambio de tipo)")
            pega.locator(SUBMIT_BUTTON).click()
            page.wait_for_timeout(STEP_PAUSE_MS)
            paso = CRONO.marcar("Actions (tras cambio de tipo)")
            pega.locator(ACTIONS_MENU).click()
        elif es_trigger:
            log.info("Tipo Miscellaneous sin Change case type en el menú: sigue con Reassign")

        # 3) Reassign. Si no aparece: MessageAttached (directo), o Step By Step, o Perform research; y después Reassign
        if not menu_option_visible(pega, REASSIGN_OPTION):
            # Un solo tiempo de espera para todas las alternativas, en orden de prioridad
            alternativas = [
                ("MessageAttached", MESSAGE_ATTACHED_OPTIONS),
                ("Step By Step", STEP_BY_STEP_OPTIONS),
                ("Perform research", (PERFORM_RESEARCH_OPTION,)),
            ]
            if not esperar_alguna_visible(pega, [sel for _, sels in alternativas for sel in sels]):
                raise TimeoutError("no apareció Reassign, MessageAttached, Step By Step ni Perform research")
            for nombre, sels in alternativas:
                sel = next((s for s in sels if pega.locator(s).count() > 0), None)
                if sel:
                    paso = CRONO.marcar(nombre)
                    pega.locator(sel).click()
                    break
            page.wait_for_timeout(STEP_PAUSE_MS)
        paso = CRONO.marcar("Reassign")
        pega.locator(REASSIGN_OPTION).click()
        page.wait_for_timeout(STEP_PAUSE_MS)

        # 4) Elegir el assignee en el dropdown "operator" (sin importar mayúsculas ni espacios)
        paso = CRONO.marcar("operator (assignee)")
        operator = pega.get_by_label(OPERATOR_LABEL)
        operator.select_option(value=valor_opcion(operator, assignee))
        page.wait_for_timeout(STEP_PAUSE_MS)

        # 5) Submit y Close (sin Save: la reasignación no deja nada por guardar)
        paso = CRONO.marcar("Submit")
        pega.locator(SUBMIT_BUTTON).click()
        page.wait_for_timeout(STEP_PAUSE_MS)
        paso = CRONO.marcar("Close")
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


def guardar_resultados(results: list, destino: Path = Path(RESULTADOS_CSV), imprimir: bool = True) -> None:
    """Escribe resultados.csv (sin las omitidas) y muestra la tabla de errores y el resumen.

    Se llama también si la corrida se corta a mitad.
    """
    out = destino
    filas = [r for r in results if r[3] != "OMITIDA"]
    with out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["fila", "assignee", "jpv", "estado", "detalle"])
        writer.writerows(filas)

    ok = [r for r in filas if r[3] == "OK"]
    errores = [r for r in filas if r[3] == "ERROR"]
    sin_procesar = [r for r in filas if r[3] == "NO PROCESADA"]
    if errores and imprimir:
        # El ancho sale del dato más largo, con 2 espacios de separación, para que el JPV no pegue con el assignee
        ancho_jpv = max(len("JPV"), *(len(r[2]) for r in errores)) + 2
        ancho_asig = max(len("Asignado a"), *(len(r[1]) for r in errores)) + 2
        print(f"\n{RED}{'JPV':<{ancho_jpv}}{'Asignado a':<{ancho_asig}}Motivo{RESET}")
        for r in errores:
            print(f"{RED}{r[2]:<{ancho_jpv}}{r[1]:<{ancho_asig}}{r[4]}{RESET}")
        print()
    log.info("Resultados: %d OK, %d con error, %d sin procesar. Detalle en %s",
             len(ok), len(errores), len(sin_procesar), out.resolve())


def procesar_filas(rows: list, failed: bool, results: list) -> None:
    """Abre un navegador, hace el login y procesa las filas. Va agregando a results."""
    with sync_playwright() as p:
        browser = p.chromium.launch(channel=BROWSER_CHANNEL, headless=HEADLESS, args=["--start-maximized"])
        page = browser.new_page(no_viewport=True)
        page.goto(APP_URL)

        if AUTO_SSO_CLICKS:
            login_sso(page)

        log.info("Completá usuario, contraseña y MFA en el navegador. El script arranca solo cuando aparezca la búsqueda.")
        page.locator(SEARCH_INPUT).wait_for(state="visible", timeout=LOGIN_TIMEOUT_MS)

        detenido = False
        for row_no, assignee, jpv, team, case_type in rows:
            if detenido:
                results.append([row_no, assignee, jpv, "NO PROCESADA", "corrida detenida por un error anterior"])
                continue
            if failed:
                # En failed, "team" es el Sub LOB (columna E). No hay equipo ni tipo de caso.
                if not any(contiene_valor(team, v) for v in FAILED_SUB_LOBS):
                    motivo = f"sub LOB '{team}' fuera de {' o '.join(FAILED_SUB_LOBS)}"
                    log.warning(f"{YELLOW}JPV %s: %s. Se saltea la fila.{RESET}", jpv, motivo)
                    results.append([row_no, assignee, jpv, "OMITIDA", motivo])
                    continue
            else:
                if team not in TEAMS_ALLOWED:
                    motivo = f"equipo '{team}' fuera de TEAMS_ALLOWED"
                    log.warning(f"{YELLOW}JPV %s: %s. Se saltea la fila.{RESET}", jpv, motivo)
                    results.append([row_no, assignee, jpv, "OMITIDA", motivo])
                    continue
                if solo_autocv(team, case_type):
                    motivo = f"tipo '{case_type}' solo se procesa para el equipo {TEAM_AUTOCV}"
                    log.warning(f"{YELLOW}JPV %s: %s. Se saltea la fila.{RESET}", jpv, motivo)
                    results.append([row_no, assignee, jpv, "OMITIDA", motivo])
                    continue
            if failed:
                owners = OWNERS_POR_SUB_LOB.get(team, OWNERS_ALLOWED)  # en failed, "team" es el Sub LOB
            else:
                owners = OWNERS_POR_EQUIPO.get(team, OWNERS_ALLOWED)
            try:
                process_row(page, assignee, jpv, case_type, owners)
                log.info(f"{GREEN}OK - JPV %s asignado a %s{RESET}", jpv, assignee)
                log.info("Tiempos JPV %s: %s", jpv, CRONO.resumen())
                results.append([row_no, assignee, jpv, "OK", ""])
            except SkipRow as skip:
                log.warning(f"{YELLOW}JPV %s: %s. Se saltea la fila.{RESET}", jpv, skip)
                results.append([row_no, assignee, jpv, "OMITIDA", str(skip)])
            except Exception as exc:  # noqa: BLE001
                text = str(exc).strip()
                msg = text.splitlines()[0] if text else type(exc).__name__
                log.error(f"{RED}ERROR - JPV %s no se pudo asignar a %s. Motivo: %s{RESET}", jpv, assignee, msg)
                log.info("Tiempos JPV %s: %s", jpv, CRONO.resumen())
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


def tramo_de_filas(rows: list, parte: int, instancias: int) -> list:
    """Parte contigua número `parte` (1 a instancias) de las filas. Todas las instancias calculan lo mismo."""
    n = len(rows)
    return rows[(parte - 1) * n // instancias: parte * n // instancias]


def unir_resultados(instancias: int) -> None:
    """Junta los resultados_parteK.csv en resultados.csv y muestra la tabla de errores y el resumen."""
    filas, faltan = [], []
    for k in range(1, instancias + 1):
        parte = Path(f"resultados_parte{k}.csv")
        if not parte.exists():
            faltan.append(k)
            continue
        with parte.open(newline="", encoding="utf-8") as f:
            lector = csv.reader(f)
            next(lector, None)  # encabezado
            filas.extend(fila for fila in lector if fila)
        parte.unlink()
    if faltan:
        log.error(f"{RED}Faltan los resultados de la(s) instancia(s): %s{RESET}", ", ".join(map(str, faltan)))
    guardar_resultados(filas, Path(RESULTADOS_CSV), imprimir=True)


def lanzar_instancias(xlsx_path: Path, rows: list, instancias: int) -> None:
    """Lanza una instancia del script por cada tramo de filas, con una pausa entre arranques, y después une los resultados."""
    script = Path(__file__).resolve()
    procesos = []
    try:
        for k in range(1, instancias + 1):
            if k > 1:
                time.sleep(ESPERA_ENTRE_INSTANCIAS_MS / 1000)
            tramo = tramo_de_filas(rows, k, instancias)
            if tramo:
                log.info("Lanzando instancia %d de %d: %d filas (de la fila %s a la %s)",
                         k, instancias, len(tramo), tramo[0][0], tramo[-1][0])
            procesos.append(subprocess.Popen(
                [sys.executable, str(script), str(xlsx_path), "--parte", str(k), "--instancias", str(instancias)]))
        for proceso in procesos:
            proceso.wait()
    except KeyboardInterrupt:
        log.error(f"{RED}Interrumpido. Se cierran las instancias.{RESET}")
        for proceso in procesos:
            proceso.terminate()
        for proceso in procesos:
            proceso.wait()
    finally:
        unir_resultados(instancias)


def main() -> None:
    parser = argparse.ArgumentParser(description="Reasigna casos en PRPC desde un XLSX.")
    parser.add_argument("xlsx", help="archivo .xlsx de entrada")
    parser.add_argument("--instancias", type=int, default=INSTANCIAS,
                        help=f"ventanas de Edge en paralelo (por defecto {INSTANCIAS})")
    parser.add_argument("--parte", type=int, help="uso interno: número de la instancia hija")
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx)
    if xlsx_path.suffix.lower() != ".xlsx":
        sys.exit("El archivo de entrada tiene que ser .xlsx")
    if args.instancias < 1:
        sys.exit("--instancias tiene que ser 1 o más")

    failed = FAILED_FILE_MARK in xlsx_path.name.casefold()
    rows = read_rows(xlsx_path, failed)
    if failed:
        log.info("Archivo failed: solo se procesan las filas con Sub LOB %s", " o ".join(FAILED_SUB_LOBS))
    log.info("%d filas para procesar", len(rows))
    if not rows:
        return

    if args.parte:  # instancia hija: procesa su tramo y deja su propio CSV
        for handler in logging.getLogger().handlers:
            handler.setFormatter(logging.Formatter(f"%(asctime)s  %(levelname)-7s [parte {args.parte}] %(message)s"))
        tramo = tramo_de_filas(rows, args.parte, args.instancias)
        results = []
        try:
            procesar_filas(tramo, failed, results)
        finally:
            guardar_resultados(results, Path(f"resultados_parte{args.parte}.csv"), imprimir=False)
        return

    if args.instancias == 1:
        results = []
        try:
            procesar_filas(rows, failed, results)
        finally:
            guardar_resultados(results)
        return

    lanzar_instancias(xlsx_path, rows, args.instancias)


if __name__ == "__main__":
    main()
