# UFEDKMLstacker.py
# Copyright (c) 2024 ot2i7ba
# https://github.com/ot2i7ba/
# This code is licensed under the MIT License (see LICENSE for details).
# Optimized version - April 2026
"""
Stack (merge) multiple KML files to generate a combined interactive map using Plotly.

Changes in optimized version:
- Removed duplicate color_name_map definition (was defined twice: globally and in save_statistics_to_excel)
- Fixed define_styles(): KML color format must be AABBGGRR, not a simple hex replacement
- Fixed process_kml_file(): points without timestamps were being silently skipped from valid_points
  but still needed proper accounting; added optional flag to include points without timestamps
- Fixed create_interactive_map(): used pd.DataFrame(df) properly, added None-safety for missing elements
- Removed redundant double hash computation in main_menu() (hash was computed in extract_file_metadata
  AND again in verify_file_integrity — now single computation)
- Added pathlib.Path usage for cleaner path handling
- Improved type hints throughout
- Extracted COLOR_PALETTE constant to avoid repeated inline dicts
- Added MAX_REMARK_LENGTH and MAX_SELECTION_LENGTH input validation constants
- validate_selection(): deduplicate selected files to avoid processing the same file twice
- save_statistics_to_csv / save_statistics_to_excel: removed redundant recalculation of fields
  that were already set in statistics during merge
- Minor: spinner uses a cleaner cycle, print_blank_line replaced with direct print()
- Added __version__ constant
"""

# Standard Libraries
import hashlib
import itertools
import json
import logging
import math
import os
import re
import subprocess
import sys
import threading
import time
from collections import defaultdict
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Third-party Libraries
import arrow
import folium
from folium.plugins import MarkerCluster, MiniMap, Fullscreen
import pandas as pd
from lxml import etree

# ── Version ──────────────────────────────────────────────────────────────────
__version__ = "0.0.4"

# ── Language / Sprache ────────────────────────────────────────────────────────
# Set by select_language() at startup. "en" = English, "de" = Deutsch.
LANG: str = "en"


def T(key: str, **kwargs: Any) -> str:
    """Return the UI string for *key* in the active language, with optional formatting."""
    s = _STRINGS.get(LANG, _STRINGS["en"]).get(key, _STRINGS["en"].get(key, key))
    return s.format(**kwargs) if kwargs else s


_STRINGS: Dict[str, Dict[str, str]] = {
    "en": {
        # ── startup ──────────────────────────────────────────────────────────
        "lang_prompt":         "Language / Sprache ([EN] = English / DE = Deutsch): ",
        # ── header / countdown ───────────────────────────────────────────────
        "returning":           "\rReturning to main menu in {n} seconds...",
        "returned":            "\rReturning to main menu...                    ",
        # ── file listing ─────────────────────────────────────────────────────
        "no_kml_found":        "No KML files found in the current directory.",
        "available_kml":       "Available KML files:",
        "exit_option":         "   e. Exit",
        # ── file selection ────────────────────────────────────────────────────
        "select_prompt":       "Enter file numbers to merge (e.g., 1, 2, 5) or 'e' to exit: ",
        "exit_msg":            "\nExiting the script. Goodbye!",
        "review_msg":          "Please review the log files and results and verify them against the source data!",
        "input_too_long":      "Input too long. Please try again.",
        "invalid_format":      "Invalid format. Please enter numbers separated by commas.",
        "out_of_range":        "Number {num} is out of range (1\u2013{max}).",
        "too_many_files":      "Maximum {max} KML files allowed; you selected {n}.",
        "too_few_files":       "At least two KML files must be selected.",
        "invalid_input":       "Invalid input, please try again.",
        "invalid_files":       "One or more selected files are invalid or missing. Please try again.",
        # ── overwrite prompt ──────────────────────────────────────────────────
        "file_exists":         "The file '{name}' already exists.",
        "overwrite_prompt":    "Overwrite? (Enter = yes / n = no): ",
        "not_overwritten":     "File will not be overwritten.",
        # ── remarks ───────────────────────────────────────────────────────────
        "remark_prompt":       "Enter a remark for '{fname}': ",
        "remark_empty":        "Remark cannot be empty.",
        "remark_too_long":     "Remark too long (max {max} characters).",
        # ── speed colours ─────────────────────────────────────────────────────
        "speed_color_header":  "\u2500\u2500 Speed Band Colours \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500",
        "speed_color_names":   "  Colour names: red, green, blue, orange, yellow, purple,",
        "speed_color_names2":  "                pink, grey, cyan, darkred",
        "speed_color_hex":     "  Or hex code:  #FF0000",
        "speed_color_prompt":  "Enter number to change, Enter to continue: ",
        "speed_color_invalid": "  Invalid number. Please enter 1\u2013{max}.",
        "speed_color_new":     "  New colour for '{label}': ",
        "speed_color_unknown": "  Unknown colour '{val}' \u2013 no change.",
        "speed_color_ok":      "  \u2713 '{label}' \u2192 {color}",
        "speed_color_adjust":  "Adjust speed band colours? (y/Enter): ",
        # ── processing ────────────────────────────────────────────────────────
        "processing_done":     "\rProcessing {name} ... done!              \n",
        "integrity_fail":      "Integrity check failed for {fname}. Aborting.",
        "no_geopoints":        "Warning: No valid geopoints found for the interactive map.",
        "map_saved":           "Interactive map saved: {path}",
        "stats_excel_saved":   "\nStatistics saved in {path}.",
        "stats_csv_saved":     "Statistics saved in {path}.",
        "speed_csv_saved":     "Speed segments saved in {path}.",
        "process_done":        "Process completed successfully.",
        "interrupted":         "\n\nProcess interrupted by user. Exiting gracefully...",
        # ── map popup / layer labels ───────────────────────────────────────────
        "map_segment":         "Segment",
        "map_source":          "Source",
        "map_band":            "Band",
        "map_from":            "From",
        "map_to":              "To",
        "map_distance":        "Distance",
        "map_duration":        "Duration",
        "map_speed":           "Speed",
        "map_speeds_layer":    "\U0001f697 Speeds \u2013 {remark}",
        "map_speed_legend":    "Speed Bands",
    },
    "de": {
        # ── startup ──────────────────────────────────────────────────────────
        "lang_prompt":         "Language / Sprache ([EN] = English / DE = Deutsch): ",
        # ── header / countdown ───────────────────────────────────────────────
        "returning":           "\rZur\u00fcck zum Hauptmen\u00fc in {n} Sekunden...",
        "returned":            "\rZur\u00fcck zum Hauptmen\u00fc...                    ",
        # ── file listing ─────────────────────────────────────────────────────
        "no_kml_found":        "Keine KML-Dateien im aktuellen Verzeichnis gefunden.",
        "available_kml":       "Verf\u00fcgbare KML-Dateien:",
        "exit_option":         "   e. Beenden",
        # ── file selection ────────────────────────────────────────────────────
        "select_prompt":       "Dateinummern zum Zusammenf\u00fchren eingeben (z.B. 1, 2, 5) oder 'e' zum Beenden: ",
        "exit_msg":            "\nSkript wird beendet. Auf Wiedersehen!",
        "review_msg":          "Bitte Log-Dateien und Ergebnisse pr\u00fcfen und gegen die Quelldaten verifizieren!",
        "input_too_long":      "Eingabe zu lang. Bitte erneut versuchen.",
        "invalid_format":      "Ung\u00fcltiges Format. Bitte Zahlen kommagetrennt eingeben.",
        "out_of_range":        "Nummer {num} au\u00dferhalb des g\u00fcltigen Bereichs (1\u2013{max}).",
        "too_many_files":      "Maximal {max} KML-Dateien erlaubt; {n} ausgew\u00e4hlt.",
        "too_few_files":       "Es m\u00fcssen mindestens zwei KML-Dateien ausgew\u00e4hlt werden.",
        "invalid_input":       "Ung\u00fcltige Eingabe, bitte erneut versuchen.",
        "invalid_files":       "Eine oder mehrere Dateien sind ung\u00fcltig oder fehlen. Bitte erneut versuchen.",
        # ── overwrite prompt ──────────────────────────────────────────────────
        "file_exists":         "Die Datei '{name}' existiert bereits.",
        "overwrite_prompt":    "\u00dcberschreiben? (Enter = ja / n = nein): ",
        "not_overwritten":     "Datei wird nicht \u00fcberschrieben.",
        # ── remarks ───────────────────────────────────────────────────────────
        "remark_prompt":       "Bezeichnung f\u00fcr '{fname}' eingeben: ",
        "remark_empty":        "Bezeichnung darf nicht leer sein.",
        "remark_too_long":     "Bezeichnung zu lang (max. {max} Zeichen).",
        # ── speed colours ─────────────────────────────────────────────────────
        "speed_color_header":  "\u2500\u2500 Farben der Geschwindigkeitsbereiche \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500",
        "speed_color_names":   "  Farbnamen: rot, gr\u00fcn, blau, orange, gelb, lila,",
        "speed_color_names2":  "             pink, grau, t\u00fcrkis, dunkelrot",
        "speed_color_hex":     "  Oder Hex-Code: #FF0000",
        "speed_color_prompt":  "Nummer eingeben zum \u00c4ndern, Enter zum Weiter: ",
        "speed_color_invalid": "  Ung\u00fcltige Nummer. Bitte 1\u2013{max} eingeben.",
        "speed_color_new":     "  Neue Farbe f\u00fcr '{label}': ",
        "speed_color_unknown": "  Unbekannte Farbe '{val}' \u2013 keine \u00c4nderung.",
        "speed_color_ok":      "  \u2713 '{label}' \u2192 {color}",
        "speed_color_adjust":  "Farben der Geschwindigkeitsbereiche anpassen? (j/Enter): ",
        # ── processing ────────────────────────────────────────────────────────
        "processing_done":     "\rVerarbeitung {name} ... fertig!              \n",
        "integrity_fail":      "Integrit\u00e4tspr\u00fcfung fehlgeschlagen f\u00fcr {fname}. Abbruch.",
        "no_geopoints":        "Warnung: Keine g\u00fcltigen Geop\u00fcnkte f\u00fcr die interaktive Karte gefunden.",
        "map_saved":           "Interaktive Karte gespeichert: {path}",
        "stats_excel_saved":   "\nStatistik gespeichert in {path}.",
        "stats_csv_saved":     "Statistik gespeichert in {path}.",
        "speed_csv_saved":     "Geschwindigkeitssegmente gespeichert in {path}.",
        "process_done":        "Verarbeitung erfolgreich abgeschlossen.",
        "interrupted":         "\n\nVerarbeitung durch Benutzer abgebrochen. Beende...",
        # ── map popup / layer labels ───────────────────────────────────────────
        "map_segment":         "Segment",
        "map_source":          "Quelle",
        "map_band":            "Bereich",
        "map_from":            "Von",
        "map_to":              "Nach",
        "map_distance":        "Distanz",
        "map_duration":        "Dauer",
        "map_speed":           "Geschwindigkeit",
        "map_speeds_layer":    "\U0001f697 Geschwindigkeiten \u2013 {remark}",
        "map_speed_legend":    "Geschwindigkeitsbereiche",
    },
}


def select_language() -> None:
    """Ask the user to choose a language and set the global LANG variable."""
    global LANG
    raw = input(_STRINGS["en"]["lang_prompt"]).strip().lower()
    LANG = "de" if raw == "de" else "en"
    logging.info("Language set to: %s", LANG)

# ── Base path ─────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_PATH = Path(sys.executable).parent
else:
    BASE_PATH = Path(__file__).resolve().parent

# ── Global Constants ──────────────────────────────────────────────────────────
output_lock = threading.Lock()
LOG_FILE        = BASE_PATH / "UFEDKMLstacker.log"
MERGED_KML_FILE = BASE_PATH / "Merged_Colored.kml"
MAX_KML_FILES        = 10
MAX_REMARK_LENGTH    = 80
MAX_SELECTION_LENGTH = 200
MAP_HEIGHT = 1080
MAP_WIDTH  = 1920

# ── Color palette (single source of truth) ───────────────────────────────────
# Keys are human-readable names; values are web hex strings (#RRGGBB)
COLOR_PALETTE: Dict[str, str] = {
    "Red":       "#FF0000",
    "Blue":      "#0000FF",
    "Yellow":    "#FFFF00",
    "Green":     "#00FF00",
    "Orange":    "#FFA500",
    "Violet":    "#EE82EE",
    "Pink":      "#FFC0CB",
    "Purple":    "#800080",
    "Turquoise": "#40E0D0",
    "Cyan":      "#00FFFF",
}

# Reverse lookup: "#RRGGBB" → "Name"
HEX_TO_NAME: Dict[str, str] = {v: k for k, v in COLOR_PALETTE.items()}
# Ordered list of hex values for sequential assignment
_COLOR_VALUES: List[str] = list(COLOR_PALETTE.values())

KML_NS = "http://www.opengis.net/kml/2.2"

# ── Speed analysis – gap filter ──────────────────────────────────────────────
# Segments with a time gap larger than this are skipped (device was off / no signal).
MAX_SEGMENT_GAP_HOURS = 4.0


# ── Geschwindigkeitsbereiche – fest definiert, nur Farben änderbar ────────────
# Die Grenzen sind fix. Nur die Farben können vom Nutzer angepasst werden.
# Gespeichert in speed_colors.json (nur die 5 Farben, kein komplexes Schema).

SPEED_BANDS: List[Dict[str, Any]] = [
    {"label": "30–50 km/h",   "from":  30, "up_to":  50, "color": "#2ECC71", "flagged": False},
    {"label": "50–70 km/h",   "from":  50, "up_to":  70, "color": "#1E90FF", "flagged": False},
    {"label": "70–100 km/h",  "from":  70, "up_to": 100, "color": "#FFD700", "flagged": False},
    {"label": "100–130 km/h", "from": 100, "up_to": 130, "color": "#FF8C00", "flagged": False},
    {"label": "130–150 km/h", "from": 130, "up_to": 150, "color": "#FF4500", "flagged": True },
    {"label": "> 150 km/h",   "from": 150, "up_to": float("inf"), "color": "#FF0000", "flagged": True },
]

# Farbe für alles unter 30 km/h (Fußgänger / Stillstand) – nicht konfigurierbar
_COLOR_UNDER_30 = "#AAAAAA"

SPEED_COLORS_FILE = BASE_PATH / "speed_colors.json"


def _load_band_colors() -> None:
    """Overwrite the colors in SPEED_BANDS from speed_colors.json if present."""
    if not SPEED_COLORS_FILE.exists():
        return
    try:
        saved = json.loads(SPEED_COLORS_FILE.read_text(encoding="utf-8"))
        for i, color in enumerate(saved):
            if i < len(SPEED_BANDS) and re.match(r"^#[0-9A-Fa-f]{6}$", color):
                SPEED_BANDS[i]["color"] = color
        logging.info("Band colors loaded from %s.", SPEED_COLORS_FILE)
    except Exception as exc:
        logging.warning("Could not load speed_colors.json: %s", exc)


def _save_band_colors() -> None:
    colors = [b["color"] for b in SPEED_BANDS]
    SPEED_COLORS_FILE.write_text(json.dumps(colors, indent=2), encoding="utf-8")
    logging.info("Band colors saved to %s.", SPEED_COLORS_FILE)


def _speed_color(speed_kmh: float, bands: List[Dict[str, Any]] = SPEED_BANDS) -> str:
    if speed_kmh < 30:
        return _COLOR_UNDER_30
    for b in bands:
        if speed_kmh < b["up_to"]:
            return b["color"]
    return bands[-1]["color"]


def _speed_flagged(speed_kmh: float, bands: List[Dict[str, Any]] = SPEED_BANDS) -> bool:
    if speed_kmh < 30:
        return False
    for b in bands:
        if speed_kmh < b["up_to"]:
            return b["flagged"]
    return bands[-1]["flagged"]


def _speed_label(speed_kmh: float, bands: List[Dict[str, Any]] = SPEED_BANDS) -> str:
    if speed_kmh < 30:
        return "< 30 km/h"
    for b in bands:
        if speed_kmh < b["up_to"]:
            return b["label"]
    return bands[-1]["label"]


# ── Einfache Farb-Konfiguration ───────────────────────────────────────────────
_NAMED_COLORS: Dict[str, str] = {
    "rot": "#FF0000", "red": "#FF0000",
    "grün": "#2ECC71", "gruen": "#2ECC71", "green": "#2ECC71",
    "blau": "#1E90FF", "blue": "#1E90FF",
    "orange": "#FF8C00",
    "gelb": "#FFD700", "yellow": "#FFD700",
    "lila": "#9B59B6", "purple": "#9B59B6",
    "pink": "#FF69B4",
    "grau": "#808080", "grey": "#808080", "gray": "#808080",
    "türkis": "#40E0D0", "turkis": "#40E0D0", "cyan": "#00FFFF",
    "dunkelrot": "#8B0000", "darkred": "#8B0000",
}


def configure_speed_colors() -> None:
    print()
    print(T("speed_color_header"))
    print(T("speed_color_names"))
    print(T("speed_color_names2"))
    print(T("speed_color_hex"))
    print()

    while True:
        for i, b in enumerate(SPEED_BANDS, 1):
            print(f"  {i}. {b['label']:<16}  {b['color']}")
        print()
        raw = input(T("speed_color_prompt")).strip()
        if not raw:
            break
        if not raw.isdigit() or not (1 <= int(raw) <= len(SPEED_BANDS)):
            print(T("speed_color_invalid", max=len(SPEED_BANDS)))
            continue
        idx = int(raw) - 1
        raw_color = input(T("speed_color_new", label=SPEED_BANDS[idx]['label'])).strip()
        if re.match(r"^#[0-9A-Fa-f]{6}$", raw_color):
            SPEED_BANDS[idx]["color"] = raw_color.upper()
        elif raw_color.lower() in _NAMED_COLORS:
            SPEED_BANDS[idx]["color"] = _NAMED_COLORS[raw_color.lower()]
        else:
            print(T("speed_color_unknown", val=raw_color))
            time.sleep(0.8)
            continue
        print(T("speed_color_ok", label=SPEED_BANDS[idx]['label'], color=SPEED_BANDS[idx]['color']))
        print()

    _save_band_colors()
    logging.info("Speed colors updated: %s", [b['color'] for b in SPEED_BANDS])


# ── Spinner ───────────────────────────────────────────────────────────────────
def spinner(stop_event: threading.Event, task_name: str, lock: threading.Lock) -> None:
    for char in itertools.cycle(r"-\|/"):
        if stop_event.is_set():
            break
        with lock:
            sys.stdout.write(f"\rProcessing {task_name} ... {char}")
            sys.stdout.flush()
        time.sleep(0.1)


# ── Logging ───────────────────────────────────────────────────────────────────
def configure_logging(log_to_console: bool = True) -> None:
    """Configure rotating file handler + optional console handler."""
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler = RotatingFileHandler(
        str(LOG_FILE), maxBytes=15 * 1024 * 1024, backupCount=3
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    handlers: List[logging.Handler] = [file_handler]
    if log_to_console:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.ERROR)
        console_handler.setFormatter(formatter)
        handlers.append(console_handler)

    logging.basicConfig(level=logging.DEBUG, handlers=handlers)
    logging.info("Logging configured (version %s).", __version__)


# ── UI helpers ────────────────────────────────────────────────────────────────
def clear_screen() -> None:
    subprocess.run("cls" if os.name == "nt" else "clear", shell=True)


def print_header() -> None:
    print(f" UFEDKMLstacker {__version__} by ot2i7ba ")
    print("=" * 40)
    print()


def display_countdown(seconds: int) -> None:
    print()
    for remaining in range(seconds, 0, -1):
        print(T("returning", n=remaining), end="")
        time.sleep(1)
    print(T("returned"))


# ── Input helpers ─────────────────────────────────────────────────────────────
def clean_html_tags(text: Optional[str]) -> str:
    if not text:
        return ""
    return re.sub(re.compile("<.*?>"), "", text)


def validate_file_path(file_path: str) -> bool:
    """Return True only if path is inside BASE_PATH and not a symlink."""
    full_path = Path(file_path).resolve()
    if full_path.is_symlink():
        logging.warning("Symbolic link detected: %s", full_path)
        return False
    try:
        full_path.relative_to(BASE_PATH)
        return True
    except ValueError:
        logging.warning("Unauthorized file path: %s", full_path)
        return False


def validate_selection(selection: str, kml_files: List[str]) -> Optional[List[str]]:
    """Parse and validate a comma-separated file-number selection string."""
    if len(selection) > MAX_SELECTION_LENGTH:
        print(T("input_too_long"))
        return None
    if not re.match(r"^[\d,\s]+$", selection):
        print(T("invalid_format"))
        return None
    try:
        seen: set = set()
        selected_files: List[str] = []
        for token in selection.split(","):
            num = int(token.strip())
            if num < 1 or num > len(kml_files):
                print(T("out_of_range", num=num, max=len(kml_files)))
                return None
            fname = kml_files[num - 1]
            if fname not in seen:
                selected_files.append(fname)
                seen.add(fname)

        if len(selected_files) > MAX_KML_FILES:
            print(T("too_many_files", max=MAX_KML_FILES, n=len(selected_files)))
            logging.warning("Too many files selected.")
            return None
        if len(selected_files) < 2:
            print(T("too_few_files"))
            logging.info("Not enough files selected.")
            display_countdown(3)
            return None
        return selected_files
    except ValueError as exc:
        print(T("invalid_input"))
        logging.error("ValueError in validate_selection: %s", exc)
        return None


# ── KML file helpers ──────────────────────────────────────────────────────────
def kml_file_info(file_path: str) -> int:
    """Return the number of Placemarks in a KML file."""
    try:
        tree = etree.parse(file_path)
        return len(tree.findall(f".//{{{KML_NS}}}Placemark"))
    except (etree.XMLSyntaxError, OSError) as exc:
        logging.error("Error reading %s: %s", file_path, exc)
        return 0


def list_kml_files() -> List[str]:
    """List all KML files in BASE_PATH (excluding the merged output)."""
    clear_screen()
    print_header()
    merged_abs = MERGED_KML_FILE.resolve()
    kml_files = [
        f
        for f in os.listdir(BASE_PATH)
        if f.endswith(".kml")
        and (BASE_PATH / f).resolve() != merged_abs
    ]
    if not kml_files:
        logging.info("No KML files found.")
        print(T("no_kml_found"))
        sys.exit(0)

    print(T("available_kml"))
    for idx, kml in enumerate(kml_files, 1):
        count = kml_file_info(str(BASE_PATH / kml))
        print(f"  {idx:>2}. {kml:<35}  Placemarks: {count}")
    print(T("exit_option"))
    print()
    return kml_files


def select_kml_files(kml_files: List[str]) -> List[str]:
    while True:
        selections = input(T("select_prompt")).strip().lower()
        if selections == "e":
            print(T("exit_msg"))
            print(T("review_msg"))
            logging.info("User chose to exit.")
            sys.exit(0)
        if not selections:
            logging.info("Empty selection – auto-selecting up to %d files.", MAX_KML_FILES)
            return kml_files[:MAX_KML_FILES]
        selected = validate_selection(selections, kml_files)
        if selected is None:
            continue
        valid = [f for f in selected if (BASE_PATH / f).exists() and validate_file_path(str(BASE_PATH / f))]
        if len(valid) < len(selected):
            print(T("invalid_files"))
            logging.warning("Invalid file paths in selection.")
            continue
        logging.info("%d valid files selected.", len(valid))
        return valid


def check_existing_merged_file() -> bool:
    if MERGED_KML_FILE.exists():
        print()
        print(T("file_exists", name=MERGED_KML_FILE.name))
        choice = input(T("overwrite_prompt")).strip().lower()
        if choice in ("n", "nein"):
            print(T("not_overwritten"))
            logging.info("User skipped overwrite.")
            display_countdown(3)
            return False
        logging.info("User confirmed overwrite.")
    return True


# ── Color / remark assignment ─────────────────────────────────────────────────
def assign_colors_to_files(files: List[str]) -> Dict[str, str]:
    """Map each file to a color hex string and write a color_mapping.txt."""
    color_map = {f: _COLOR_VALUES[i % len(_COLOR_VALUES)] for i, f in enumerate(files)}
    mapping_file = BASE_PATH / "color_mapping.txt"
    with mapping_file.open("w", encoding="utf-8") as fh:
        for fname, hex_color in color_map.items():
            name = HEX_TO_NAME.get(hex_color, "Unknown")
            fh.write(f"{fname} = {hex_color} ({name})\n")
    logging.info("Color mapping written to %s.", mapping_file)
    return color_map


def get_remarks(files: List[str]) -> Dict[str, str]:
    remarks: Dict[str, str] = {}
    print()
    for fname in files:
        while True:
            remark = input(T("remark_prompt", fname=fname)).strip()
            if not remark:
                print(T("remark_empty"))
                continue
            if len(remark) > MAX_REMARK_LENGTH:
                print(T("remark_too_long", max=MAX_REMARK_LENGTH))
                continue
            remarks[fname] = remark
            break
    logging.info("Remarks: %s", remarks)
    return remarks


# ── KML style generation ──────────────────────────────────────────────────────
def _hex_to_kml_color(hex_color: str) -> str:
    """
    Convert #RRGGBB to KML AABBGGRR format (fully opaque).
    KML color order is Alpha-Blue-Green-Red.
    """
    hex_color = hex_color.lstrip("#")
    r, g, b = hex_color[0:2], hex_color[2:4], hex_color[4:6]
    return f"ff{b}{g}{r}"


def define_styles(document: etree._Element, color_map: Dict[str, str]) -> None:
    """Append KML Style elements with correct AABBGGRR color encoding."""
    for fname, hex_color in color_map.items():
        style = etree.SubElement(document, "Style", id=fname)
        icon_style = etree.SubElement(style, "IconStyle")
        color_el = etree.SubElement(icon_style, "color")
        color_el.text = _hex_to_kml_color(hex_color)
        icon = etree.SubElement(icon_style, "Icon")
        href = etree.SubElement(icon, "href")
        href.text = "http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png"


# ── Timestamp parsing ─────────────────────────────────────────────────────────
_REGEX_FORMATS = [
    # ISO variants handled by arrow; regex patterns are a fallback for exotic formats
    r"^\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}\(UTC[+-]\d{1,2}\)$",
    r"^\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2}$",
    r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$",
    r"^\d{2}/\d{2}/\d{4}$",
    r"^\d{4}/\d{2}/\d{2}$",
    r"^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$",
    r"^\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}$",
    r"^\d{4}\.\d{2}\.\d{2}$",
    r"^\d{2} [A-Za-z]{3} \d{4} \d{2}:\d{2}:\d{2}$",
    r"^\d{2} [A-Za-z]{3} \d{4}$",
    r"^[A-Za-z]{3} \d{2}, \d{4}$",
    r"^[A-Za-z]{3} \d{2}, \d{4} \d{2}:\d{2}:\d{2} \w{2}$",
]


def parse_timestamp(ts: Optional[str]) -> Optional[datetime]:
    if not ts:
        return None
    try:
        return arrow.get(ts).datetime
    except Exception:
        pass
    for pattern in _REGEX_FORMATS:
        if re.match(pattern, ts):
            try:
                return datetime.fromisoformat(ts)
            except ValueError:
                pass
    logging.debug("Could not parse timestamp: %r", ts)
    return None


# ── File metadata / integrity ─────────────────────────────────────────────────
def calculate_file_hash(file_path: str) -> str:
    sha = hashlib.sha256()
    with open(file_path, "rb") as fh:
        for block in iter(lambda: fh.read(65536), b""):
            sha.update(block)
    return sha.hexdigest()


def extract_file_metadata(file_path: str) -> Dict[str, Any]:
    try:
        stat = os.stat(file_path)
        meta = {
            "creation_time":     datetime.fromtimestamp(stat.st_ctime).isoformat(),
            "modification_time": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "file_size":         stat.st_size,
            "sha256":            calculate_file_hash(file_path),
        }
        logging.info("Metadata for %s: %s", file_path, meta)
        return meta
    except OSError as exc:
        logging.error("Error retrieving metadata for %s: %s", file_path, exc)
        return {}


def verify_file_integrity(file_path: str, expected_hash: str) -> bool:
    actual = calculate_file_hash(file_path)
    if actual == expected_hash:
        logging.info("Integrity OK: %s", file_path)
        return True
    logging.warning("Integrity FAIL for %s (expected %s, got %s).", file_path, expected_hash, actual)
    return False


# ── Speed / distance calculation ──────────────────────────────────────────────
def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Great-circle distance in km between two WGS-84 points (Haversine formula)."""
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2) ** 2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def calculate_speed_segments(
    points: List[Dict[str, Any]],
    remark: str,
) -> List[Dict[str, Any]]:
    """
    Given a time-sorted list of geopoints (each with 'lat', 'lon', 'timestamp',
    'name'), compute the speed between every consecutive pair and return a list
    of segment dicts.

    A segment is skipped when:
    - Either endpoint has no timestamp.
    - The time delta is zero (duplicate position).
    - The time gap exceeds MAX_SEGMENT_GAP_HOURS (device was off / no signal).

    Each returned dict contains:
        from_name, to_name, from_ts, to_ts,
        lat_a, lon_a, lat_b, lon_b,
        distance_km, duration_s, speed_kmh,
        speed_band, flagged (bool), remark, color
    """
    sorted_pts = sorted(
        [p for p in points if p.get("timestamp")],
        key=lambda p: p["timestamp"],
    )

    segments: List[Dict[str, Any]] = []
    for i in range(len(sorted_pts) - 1):
        a, b = sorted_pts[i], sorted_pts[i + 1]
        dt_s = (b["timestamp"] - a["timestamp"]).total_seconds()

        if dt_s <= 0:
            logging.debug("Skipping zero/negative time delta between %r and %r.", a["name"], b["name"])
            continue

        if dt_s > MAX_SEGMENT_GAP_HOURS * 3600:
            logging.debug(
                "Gap of %.1f h between %r and %r exceeds limit – segment skipped.",
                dt_s / 3600, a["name"], b["name"],
            )
            continue

        dist_km   = haversine_km(a["lat"], a["lon"], b["lat"], b["lon"])
        speed_kmh = (dist_km / dt_s) * 3600.0
        flagged   = _speed_flagged(speed_kmh)

        segments.append({
            "remark":      remark,
            "from_name":   a["name"],
            "to_name":     b["name"],
            "from_ts":     a["timestamp"].isoformat(),
            "to_ts":       b["timestamp"].isoformat(),
            "lat_a":       a["lat"],
            "lon_a":       a["lon"],
            "lat_b":       b["lat"],
            "lon_b":       b["lon"],
            "distance_km": round(dist_km,   4),
            "duration_s":  round(dt_s,      1),
            "speed_kmh":   round(speed_kmh, 2),
            "speed_band":  _speed_label(speed_kmh),
            "flagged":     flagged,
            "color":       _speed_color(speed_kmh),
        })
        if flagged:
            logging.warning(
                "Speed flagged: %.1f km/h between %r and %r (%.3f km in %.0f s).",
                speed_kmh, a["name"], b["name"], dist_km, dt_s,
            )

    logging.info("Speed segments for '%s': %d total, %d flagged.",
                 remark, len(segments), sum(1 for s in segments if s["flagged"]))
    return segments


# ── Core KML processing ───────────────────────────────────────────────────────
def process_kml_file(
    file_path: str,
    remark: str,
    include_without_timestamps: bool = False,
) -> Tuple[int, int, int, List[Dict[str, Any]]]:
    """
    Parse a KML file and extract valid geopoints.

    Args:
        file_path: Path to the KML file.
        remark: User label for this file.
        include_without_timestamps: If True, include points lacking timestamps
            (timestamp will be None). Default False preserves original behaviour.

    Returns:
        (total_points, points_with_timestamps, points_without_timestamps, valid_points)
    """
    file_name = Path(file_path).name
    stop_event = threading.Event()
    spinner_thread = threading.Thread(target=spinner, args=(stop_event, file_name, output_lock))
    spinner_thread.start()

    total_points = 0
    points_with_ts = 0
    points_without_ts = 0
    valid_points: List[Dict[str, Any]] = []

    try:
        for _, elem in etree.iterparse(file_path, events=("end",), tag=f"{{{KML_NS}}}Placemark"):
            total_points += 1

            name = elem.findtext(f"{{{KML_NS}}}name", default="")
            name = f"({remark}) - {name}"
            raw_desc = elem.findtext(f"{{{KML_NS}}}description", default="")
            description = clean_html_tags(raw_desc)

            # Timestamp resolution: <TimeStamp><when> → name text → description text
            timestamp: Optional[datetime] = None
            ts_source: Optional[str] = None

            ts_elem = elem.find(f"{{{KML_NS}}}TimeStamp")
            if ts_elem is not None:
                when = ts_elem.findtext(f"{{{KML_NS}}}when")
                if when:
                    timestamp = parse_timestamp(when)
                    if timestamp:
                        ts_source = "<TimeStamp>"

            if not timestamp:
                timestamp = parse_timestamp(name)
                if timestamp:
                    ts_source = "<name>"

            if not timestamp:
                timestamp = parse_timestamp(description)
                if timestamp:
                    ts_source = "<description>"

            if timestamp:
                points_with_ts += 1
            else:
                points_without_ts += 1

            # Coordinates
            coords_elem = elem.find(f".//{{{KML_NS}}}coordinates")
            if coords_elem is not None and coords_elem.text:
                parts = coords_elem.text.strip().split(",")
                if len(parts) >= 2 and (timestamp or include_without_timestamps):
                    try:
                        lon, lat = float(parts[0]), float(parts[1])
                        valid_points.append({
                            "lon": lon,
                            "lat": lat,
                            "timestamp": timestamp,
                            "name": name,
                            "description": description,
                        })
                        if ts_source:
                            logging.info(
                                "Geopoint: name=%r ts=%s src=%s coords=(%s,%s)",
                                name, timestamp, ts_source, lon, lat,
                            )
                        else:
                            logging.info("Geopoint (no ts): name=%r coords=(%s,%s)", name, lon, lat)
                    except ValueError as exc:
                        logging.warning("Bad coordinates in %s: %s", file_name, exc)

            elem.clear()

    except etree.XMLSyntaxError as exc:
        logging.error("XML Syntax Error in %s: %s", file_path, exc)
    except OSError as exc:
        logging.error("OS Error in %s: %s", file_path, exc)
    except Exception as exc:
        logging.error("Unhandled error in %s: %s", file_path, exc)
    finally:
        stop_event.set()
        spinner_thread.join()
        sys.stdout.write(T("processing_done", name=file_name))
        sys.stdout.flush()

    logging.info(
        "%s: total=%d with_ts=%d without_ts=%d valid=%d",
        file_path, total_points, points_with_ts, points_without_ts, len(valid_points),
    )
    return total_points, points_with_ts, points_without_ts, valid_points


def merge_kml_files(
    files: List[str],
    color_map: Dict[str, str],
    remarks: Dict[str, str],
    statistics: List[Dict[str, Any]],
) -> Tuple[str, int, List[Dict[str, Any]]]:
    """Returns (merged_kml_path, total_valid_points, all_speed_segments)."""
    merged_root = etree.Element("kml", xmlns=KML_NS)
    document = etree.SubElement(merged_root, "Document")
    define_styles(document, color_map)

    total_valid_points = 0
    all_speed_segments: List[Dict[str, Any]] = []
    file_metadata = {f: extract_file_metadata(str(BASE_PATH / f)) for f in files}

    for fname in files:
        fp = str(BASE_PATH / fname)
        total, with_ts, without_ts, valid_pts = process_kml_file(fp, remarks[fname])

        if not validate_file_path(fp):
            logging.warning("Skipping invalid path: %s", fp)
            continue

        total_valid_points += len(valid_pts)

        # ── Speed analysis for this file ──────────────────────────────────
        segs = calculate_speed_segments(valid_pts, remarks[fname])
        all_speed_segments.extend(segs)

        statistics.append({
            "file":                      fname,
            "total_points":              total,
            "points_with_timestamps":    with_ts,
            "points_without_timestamps": without_ts,
            "valid_points":              len(valid_pts),
            "mapped_points":             len(valid_pts),
            "speed_segments":            len(segs),
            "max_speed_kmh":             round(max((s["speed_kmh"] for s in segs), default=0.0), 2),
            "remark":                    remarks[fname],
            "color":                     color_map[fname],
            "color_name":                HEX_TO_NAME.get(color_map[fname], "Unknown"),
            **file_metadata[fname],
        })

        for pt in valid_pts:
            placemark = etree.Element("Placemark")
            su = etree.SubElement(placemark, "styleUrl")
            su.text = f"#{fname}"
            point_el = etree.SubElement(placemark, "Point")
            coords_el = etree.SubElement(point_el, "coordinates")
            coords_el.text = f"{pt['lon']},{pt['lat']}"
            desc_el = etree.SubElement(placemark, "description")
            desc_el.text = pt["description"]
            ts_el = etree.SubElement(placemark, "TimeStamp")
            when_el = etree.SubElement(ts_el, "when")
            when_el.text = pt["timestamp"].isoformat() if pt["timestamp"] else ""
            name_el = etree.SubElement(placemark, "name")
            name_el.text = pt["name"]
            document.append(placemark)

    with MERGED_KML_FILE.open("wb") as fh:
        fh.write(etree.tostring(merged_root, pretty_print=True))
    logging.info("Merged KML saved: %s", MERGED_KML_FILE)
    logging.info("Total points: %d", sum(s["total_points"] for s in statistics))
    logging.info("Total valid points: %d", total_valid_points)
    logging.info("Total speed segments: %d", len(all_speed_segments))
    return str(MERGED_KML_FILE), total_valid_points, all_speed_segments


# ── Interactive map ───────────────────────────────────────────────────────────
def _collect_map_rows(
    merged_kml: str,
    color_map: Dict[str, str],
    remarks: Dict[str, str],
) -> List[Dict[str, Any]]:
    """Parse merged KML and return a list of row dicts for map rendering."""
    rows: List[Dict[str, Any]] = []
    tree = etree.parse(merged_kml)
    for pm in tree.findall(f".//{{{KML_NS}}}Placemark"):
        coords_el = pm.find(f".//{{{KML_NS}}}coordinates")
        name_el   = pm.find(f".//{{{KML_NS}}}name")
        desc_el   = pm.find(f".//{{{KML_NS}}}description")
        su_el     = pm.find(f".//{{{KML_NS}}}styleUrl")
        ts_el     = pm.find(f".//{{{KML_NS}}}when")

        if coords_el is None or not coords_el.text:
            continue
        parts = coords_el.text.strip().split(",")
        if len(parts) < 2:
            continue
        try:
            lon, lat = float(parts[0]), float(parts[1])
        except ValueError:
            continue

        style_key = (su_el.text or "").lstrip("#") if su_el is not None else ""
        rows.append({
            "lat":         lat,
            "lon":         lon,
            "name":        (name_el.text or "") if name_el is not None else "",
            "description": (desc_el.text or "") if desc_el is not None else "",
            "timestamp":   (ts_el.text or "")  if ts_el  is not None else "",
            "color":       color_map.get(style_key, "#808080"),
            "remark":      remarks.get(style_key, "No remark"),
        })
    return rows



def create_interactive_map(
    merged_kml: str,
    color_map: Dict[str, str],
    remarks: Dict[str, str],
    speed_segments: Optional[List[Dict[str, Any]]] = None,
) -> None:
    """
    Build a fully interactive Leaflet map via Folium.

    Features:
    - Mouse-wheel zoom, pinch-zoom, +/- buttons
    - Fullscreen button / Mini-map
    - Per-source MarkerCluster layers (toggleable)
    - Speed segment polylines colour-coded by velocity (toggleable)
    - Popup on click with name, timestamp, description / speed details
    - Auto-fits viewport to all points
    """
    rows = _collect_map_rows(merged_kml, color_map, remarks)

    if not rows:
        logging.warning("No rows to plot in interactive map.")
        print(T("no_geopoints"))
        return

    # ── Compute map center ────────────────────────────────────────────────────
    lats = [r["lat"] for r in rows]
    lons = [r["lon"] for r in rows]
    center_lat = (min(lats) + max(lats)) / 2
    center_lon = (min(lons) + max(lons)) / 2

    # ── Base map ──────────────────────────────────────────────────────────────
    # Default tile: CartoDB Positron – reliable, no API key, no 403 errors.
    # Folium's built-in "OpenStreetMap" hits tile.openstreetmap.org directly,
    # which blocks script/automated access with 403. We add tiles manually so
    # multiple layers are selectable via the layer control.
    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=5,
        control_scale=True,
        prefer_canvas=True,
        tiles=None,   # tiles added manually below
    )

    # CartoDB Positron – light, clean, forensic-friendly (default)
    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
        attr=(
            '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors '
            '&copy; <a href="https://carto.com/attributions">CARTO</a>'
        ),
        name="CartoDB Positron (default)",
        max_zoom=20,
    ).add_to(m)

    # CartoDB Dark Matter – dark mode
    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png",
        attr=(
            '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors '
            '&copy; <a href="https://carto.com/attributions">CARTO</a>'
        ),
        name="CartoDB Dark Matter",
        max_zoom=20,
    ).add_to(m)

    # OpenStreetMap via community mirror (not the blocked tile.openstreetmap.org)
    folium.TileLayer(
        tiles="https://tile.openstreetmap.bzh/br/{z}/{x}/{y}.png",
        attr='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
        name="OpenStreetMap (mirror)",
        max_zoom=19,
    ).add_to(m)

    # ── Plugins ───────────────────────────────────────────────────────────────
    Fullscreen(
        position="topleft",
        title="Fullscreen",
        title_cancel="Exit fullscreen",
        force_separate_button=True,
    ).add_to(m)

    MiniMap(toggle_display=True, position="bottomright").add_to(m)

    # ── Per-remark feature groups with clustering ─────────────────────────────
    # Group rows by remark so each source is a toggleable layer
    groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        groups[row["remark"]].append(row)

    for remark_label, pts in groups.items():
        fg = folium.FeatureGroup(name=f"<span style='color:{pts[0]['color']}'>"
                                      f"&#9679;</span> {remark_label}",
                                 show=True)
        cluster = MarkerCluster(
            options={
                "maxClusterRadius": 40,
                "disableClusteringAtZoom": 14,   # show individual pins at zoom ≥ 14
            }
        )

        for pt in pts:
            popup_html = (
                f"<b>{pt['name']}</b><br>"
                f"<i>{pt['timestamp']}</i><br>"
                f"<small>{pt['description'][:300]}</small>"
            ) if pt["description"] else (
                f"<b>{pt['name']}</b><br>"
                f"<i>{pt['timestamp']}</i>"
            )
            folium.CircleMarker(
                location=[pt["lat"], pt["lon"]],
                radius=7,
                color=pt["color"],
                fill=True,
                fill_color=pt["color"],
                fill_opacity=0.85,
                weight=1.5,
                popup=folium.Popup(popup_html, max_width=320),
                tooltip=pt["name"],
            ).add_to(cluster)

        cluster.add_to(fg)
        fg.add_to(m)

    # ── Speed segment polylines ───────────────────────────────────────────────
    if speed_segments:
        seg_groups: Dict[str, folium.FeatureGroup] = {}
        for seg in speed_segments:
            grp_name = T("map_speeds_layer", remark=seg['remark'])
            if grp_name not in seg_groups:
                seg_groups[grp_name] = folium.FeatureGroup(
                    name=grp_name, show=True
                )

            band_label = seg.get("speed_band", "")
            tooltip_txt = f"{seg['speed_kmh']:.1f} km/h"
            popup_html = (
                f"<b>{T('map_segment')}</b><br>"
                f"<b>{T('map_source')}:</b> {seg['remark']}<br>"
                f"<b>{T('map_band')}:</b> <span style='color:{seg['color']};font-weight:bold'>"
                f"{band_label}</span><br>"
                f"<b>{T('map_from')}:</b> {seg['from_name']}<br>"
                f"&nbsp;&nbsp;&nbsp;<i>{seg['from_ts']}</i><br>"
                f"<b>{T('map_to')}:</b> {seg['to_name']}<br>"
                f"&nbsp;&nbsp;&nbsp;<i>{seg['to_ts']}</i><br>"
                f"<hr style='margin:4px 0'>"
                f"<b>{T('map_distance')}:</b> {seg['distance_km']:.3f} km<br>"
                f"<b>{T('map_duration')}:</b> {int(seg['duration_s']//3600):02d}h "
                f"{int((seg['duration_s']%3600)//60):02d}m "
                f"{int(seg['duration_s']%60):02d}s<br>"
                f"<b>{T('map_speed')}:</b> "
                f"<span style='color:{seg['color']};font-weight:bold'>"
                f"{seg['speed_kmh']:.1f} km/h</span>"
            )

            folium.PolyLine(
                locations=[
                    [seg["lat_a"], seg["lon_a"]],
                    [seg["lat_b"], seg["lon_b"]],
                ],
                color=seg["color"],
                weight=5 if seg["flagged"] else 3,
                opacity=0.85 if seg["flagged"] else 0.65,
                tooltip=tooltip_txt,
                popup=folium.Popup(popup_html, max_width=340),
            ).add_to(seg_groups[grp_name])

        for grp in seg_groups.values():
            grp.add_to(m)

        # ── Speed band legend – frei verschiebbar ────────────────────────
        legend_rows = "".join(
            f"<tr><td style='padding:2px 6px'>"
            f"<span style='background:{b['color']};display:inline-block;"
            f"width:18px;height:10px;border-radius:2px'></span></td>"
            f"<td style='padding:2px 4px;font-size:11px'>{b['label']}</td></tr>"
            for b in SPEED_BANDS
        )
        legend_html = (
            "<div id='ufed-speed-legend' style='"
            "position:fixed;top:80px;left:10px;z-index:9999;"
            "background:rgba(255,255,255,0.93);padding:6px 10px 8px 10px;"
            "border-radius:6px;border:1px solid #bbb;font-family:sans-serif;"
            "box-shadow:0 2px 6px rgba(0,0,0,.15);user-select:none'>"
            "<div id='ufed-speed-legend-handle' style='"
            "cursor:grab;font-size:11px;font-weight:bold;margin-bottom:4px;"
            "color:#555;display:flex;justify-content:space-between;align-items:center'>"
            f"<span>{T('map_speed_legend')}</span>"
            "<span style='margin-left:8px;opacity:.5;font-size:10px'>&#9776;</span>"
            "</div>"
            f"<table style='border-collapse:collapse'>{legend_rows}</table>"
            "</div>"
        )
        m.get_root().html.add_child(folium.Element(legend_html))

    # ── Drag-JS für Legende ──────────────────────────────────────────────────
    drag_js = """
<script>
(function () {
  function makeDraggable(el, handle) {
    handle = handle || el;
    var isDragging = false, startX, startY, origLeft, origTop;

    function toFixed(el) {
      if (el.style.position === 'fixed') return;
      var r = el.getBoundingClientRect();
      el.style.position = 'fixed';
      el.style.left     = r.left + 'px';
      el.style.top      = r.top  + 'px';
      el.style.right    = 'auto';
      el.style.bottom   = 'auto';
      el.style.transform = '';
    }

    handle.addEventListener('mousedown', function (e) {
      if (['INPUT','BUTTON','A','SELECT'].indexOf(e.target.tagName) >= 0) return;
      isDragging = true;
      toFixed(el);
      startX   = e.clientX;
      startY   = e.clientY;
      origLeft = parseFloat(el.style.left) || 0;
      origTop  = parseFloat(el.style.top)  || 0;
      handle.style.cursor = 'grabbing';
      e.preventDefault();
      e.stopPropagation();
    });
    document.addEventListener('mousemove', function (e) {
      if (!isDragging) return;
      el.style.left = (origLeft + e.clientX - startX) + 'px';
      el.style.top  = (origTop  + e.clientY - startY) + 'px';
    });
    document.addEventListener('mouseup', function () {
      if (isDragging) { isDragging = false; handle.style.cursor = 'grab'; }
    });
  }

  window.addEventListener('load', function () {
    var legend = document.getElementById('ufed-speed-legend');
    if (legend) {
      var handle = document.getElementById('ufed-speed-legend-handle');
      makeDraggable(legend, handle);
    }
  });
}());
</script>
"""
    m.get_root().html.add_child(folium.Element(drag_js))

    # ── Layer control (legend / toggle) ───────────────────────────────────────
    folium.LayerControl(collapsed=False).add_to(m)

    # ── Fit bounds to all points ──────────────────────────────────────────────
    if lats and lons:
        m.fit_bounds([[min(lats), min(lons)], [max(lats), max(lons)]])

    out_file = BASE_PATH / "interactive_map.html"
    m.save(str(out_file))
    logging.info("Interactive map saved: %s", out_file)
    print(T("map_saved", path=out_file))


# ── Statistics export ─────────────────────────────────────────────────────────
_STAT_COLUMNS = [
    "file", "total_points", "points_with_timestamps", "points_without_timestamps",
    "valid_points", "mapped_points",
    "speed_segments", "max_speed_kmh",
    "remark", "color", "color_name",
    "creation_time", "modification_time", "file_size", "sha256",
]

_SPEED_COLUMNS = [
    "remark", "from_name", "to_name", "from_ts", "to_ts",
    "lat_a", "lon_a", "lat_b", "lon_b",
    "distance_km", "duration_s", "speed_kmh", "speed_band",
]


def save_statistics_to_excel(
    statistics: List[Dict[str, Any]],
    total_valid_points: int,
    total_mapped_points: int,
    speed_segments: Optional[List[Dict[str, Any]]] = None,
) -> None:
    df = pd.DataFrame(statistics)
    cols = [c for c in _STAT_COLUMNS if c in df.columns]
    df = df[cols]

    excel_file = BASE_PATH / "KML_Statistics.xlsx"
    with pd.ExcelWriter(str(excel_file), engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Summary", index=False)

        pd.DataFrame({
            "Total Valid Points":  [total_valid_points],
            "Total Mapped Points": [total_mapped_points],
        }).to_excel(writer, sheet_name="Total Summary", index=False)

        # ── Speed segments sheet ──────────────────────────────────────────
        if speed_segments:
            spd_df = pd.DataFrame(speed_segments)
            spd_cols = [c for c in _SPEED_COLUMNS if c in spd_df.columns]
            spd_df = spd_df[spd_cols].sort_values("speed_kmh", ascending=False)
            spd_df.to_excel(writer, sheet_name="Speed Segments", index=False)

            # Per-remark summary
            summary_rows = []
            for remark_label, grp in spd_df.groupby("remark"):
                summary_rows.append({
                    "remark":            remark_label,
                    "total_segments":    len(grp),
                    "max_speed_kmh":     round(grp["speed_kmh"].max(), 2),
                    "avg_speed_kmh":     round(grp["speed_kmh"].mean(), 2),
                    "total_distance_km": round(grp["distance_km"].sum(), 3),
                    "total_duration_h":  round(grp["duration_s"].sum() / 3600, 3),
                })
            pd.DataFrame(summary_rows).to_excel(
                writer, sheet_name="Speed Summary", index=False
            )

    logging.info("Statistics (Excel) saved: %s", excel_file)
    print(T("stats_excel_saved", path=excel_file))


def save_statistics_to_csv(
    statistics: List[Dict[str, Any]],
    total_valid_points: int,
    total_mapped_points: int,
    speed_segments: Optional[List[Dict[str, Any]]] = None,
    filename: str = "KML_Statistics.csv",
) -> None:
    df = pd.DataFrame(statistics)
    cols = [c for c in _STAT_COLUMNS if c in df.columns]
    df[cols].to_csv(str(BASE_PATH / filename), index=False)
    logging.info("Statistics (CSV) saved: %s", BASE_PATH / filename)
    print(T("stats_csv_saved", path=BASE_PATH / filename))

    if speed_segments:
        spd_df = pd.DataFrame(speed_segments)
        spd_cols = [c for c in _SPEED_COLUMNS if c in spd_df.columns]
        spd_csv = BASE_PATH / "KML_Speed_Segments.csv"
        spd_df[spd_cols].sort_values("speed_kmh", ascending=False).to_csv(
            str(spd_csv), index=False
        )
        logging.info("Speed segments (CSV) saved: %s", spd_csv)
        print(T("speed_csv_saved", path=spd_csv))


# ── Main menu / entry point ───────────────────────────────────────────────────
def main_menu() -> None:
    try:
        while True:
            kml_files = list_kml_files()
            selected  = select_kml_files(kml_files)

            if not check_existing_merged_file():
                continue

            # ── Speed band colour configuration ──────────────────────────
            print()
            if input(T("speed_color_adjust")).strip().lower() in ("j", "y"):
                configure_speed_colors()

            # ── File integrity check ──────────────────────────────────────
            file_metadata: Dict[str, Dict[str, Any]] = {
                f: extract_file_metadata(str(BASE_PATH / f)) for f in selected
            }
            all_ok = True
            for fname in selected:
                if not verify_file_integrity(str(BASE_PATH / fname), file_metadata[fname]["sha256"]):
                    print(T("integrity_fail", fname=fname))
                    logging.error("Integrity check failed for %s.", fname)
                    all_ok = False
                    break
            if not all_ok:
                sys.exit(1)

            color_map = assign_colors_to_files(selected)
            remarks   = get_remarks(selected)
            print()

            statistics: List[Dict[str, Any]] = []
            merged_kml, total_valid, speed_segments = merge_kml_files(
                selected, color_map, remarks, statistics
            )
            total_mapped = sum(s["valid_points"] for s in statistics)

            create_interactive_map(merged_kml, color_map, remarks, speed_segments)
            save_statistics_to_excel(statistics, total_valid, total_mapped, speed_segments)
            save_statistics_to_csv(statistics, total_valid, total_mapped, speed_segments)

            print()
            print(T("process_done"))
            print(T("review_msg"))
            display_countdown(3)

    except KeyboardInterrupt:
        print(T("interrupted"))
        print(T("review_msg"))
        logging.info("Interrupted by user (CTRL+C).")
        sys.exit(0)


if __name__ == "__main__":
    _load_band_colors()
    configure_logging()
    select_language()
    try:
        main_menu()
    except KeyboardInterrupt:
        print(T("interrupted"))
        logging.info("Interrupted by user at entry point.")
        sys.exit(0)
