"""UI text constants for the desktop GUI."""

from __future__ import annotations

APP_TITLE = "Automazione BDAP"
WINDOW_GEOMETRY = "1050x810"

STATUS_SELECT_DATA = "Selezionare i dati necessari"
STATUS_WORKSPACE_NOT_FOUND = "Workspace non trovato"
STATUS_READY = "Pronto"
STATUS_RUNNING = "Esecuzione in corso..."
STATUS_DONE = "Completato"
STATUS_ERROR = "Errore"
STATUS_DROP_INVALID = "Drop non valido: trascina una cartella"
STATUS_LOAD_COMUNI_FIRST = "Carica prima una cartella dati valida e aggiorna i comuni."
PROGRESS_PREPARING = "Preparazione file..."
PROGRESS_DONE = "Compilazione completata"
PROGRESS_LABEL_TEMPLATE = "{percent:.0f}% - {message}"

COMMON_COUNT_PREFIX = "Comuni trovati"
ALL_YEARS_LABEL = "Tutti i presenti"

HEADER_TITLE = "Automazione BDAP"
HEADER_SUBTITLE = "Seleziona un comune e compila il template con i dati BDAP disponibili."

CARD_LOADING = "Caricamento Dati"
FIELD_WORKSPACE_ROOT = "Workspace root"
BUTTON_BROWSE_FOLDER = "Sfoglia cartella"
BUTTON_REFRESH_COMUNI = "Aggiorna comuni"
DROP_ZONE_TEXT = "Trascina qui la cartella dati completa"
DROP_ZONE_TEXT_NO_DND = "Drag and drop non disponibile: usa 'Sfoglia cartella'."

CARD_PARAMETERS = "Parametri"
FIELD_COMUNE = "Seleziona il comune *"
FIELD_YEARS = "Seleziona anni"
HINT_YEAR_RANGE = "Oppure seleziona intervallo anni"
FIELD_YEAR_FROM = "Anno da"
FIELD_YEAR_TO = "Anno a"
BUTTON_RUN = "Genera Report Excel"

ERROR_TITLE_INPUT = "Errore input"
ERROR_TITLE_GENERIC = "Errore"
DONE_TITLE = "Completato"
DONE_MESSAGE_TEMPLATE = "Elaborazione terminata per il comune '{comune}'."
DONE_MESSAGE_WITH_CONTROLLI_TEMPLATE = (
    "Elaborazione terminata per il comune '{comune}'.\n\n"
    "{controlli_summary}"
)

NO_COMUNI_FOUND_MESSAGE = (
    "Nessun comune trovato: seleziona una root comuni, una cartella comune, "
    "1 doc/dati BDAP o una cartella con file BDAP .xlsx contenenti anno "
    "(YYYY oppure YY in nomi tipo Rend/BDAP). Evita nomi ambigui con piu YY."
)

CELL_COMMENT_REVIEW_OR_QUESTIONNAIRE = "TODO: revisione o questionario"
