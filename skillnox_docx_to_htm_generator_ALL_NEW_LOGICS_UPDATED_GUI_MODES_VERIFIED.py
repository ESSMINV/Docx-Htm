#!/usr/bin/env python3
"""
SkillNox DOCX -> self-contained HTM generator

Designed for the 15-column SkillNox question-bank DOCX schema used by the
attached Kinematics source. The generated HTM keeps the complete quiz/test,
palette, zoom, theme, SVG, MathJax, responsive and two-column logic from the
latest user-approved HTM master shell, including all current desktop/mobile, template, palette, zoom, SVG, CRM hierarchy, CRM radial controls, touch-board ESVG gestures, Case Study, and test-review logic.

Windows GUI use:
    1. Put the DOCX beside two folders named imagesQ and imagesE.
    2. Double-click this .py file (or run it with Python).
    3. Select the DOCX in the graphical window.
    4. Review or add Qtype (P/T) code/full-form mappings. Custom Qtypes are saved automatically.
    5. Review or add Question Type (QTYPE column) code/full-form mappings such as MCQ, BLNK, and MTCH.
    6. Review the detected total scored-question count and configure Question Attempts Max, Default Test Attempts Max, and Test Attempt Rules.
    7. Select one or more HTM modes: Self Learning, Test, and CRM.
       - All three selected: the normal full three-mode HTM is generated.
       - A subset selected: only those modes are available in the generated HTM.
       - If Self Learning is not selected, the HTM opens directly in Test (preferred) or CRM.
    8. Click Generate HTM & Open Chrome. The finished file opens automatically.

Command-line use:
    python skillnox_docx_to_htm.py input.docx -o output.htm

No third-party Python package is required. DOCX is read directly as OOXML.

SVG comparison is intentionally disabled. Every SVG referenced by the DOCX
is read from the selected imagesQ/imagesE folders and inlined into the output.
Changed SVG XML is accepted; only genuinely missing or unreadable SVG files
stop generation.

The bundled output shell is the latest validated no-Base64 consolidated HTM, including immediate startup shimmer, Test Setup/filter colours and live counts, stable viewport-fitted ESVG, explanation-render fail-safe, repeated-test reset, main-banner removal, rail-based Eye control, Slide card/top-border fixes, independent per-slide smart scrolling with native browser sensitivity, Slide+Question-Palette sticky-left ESVG behavior through the Explanation boundary, a Question Palette that keeps the Menu button visible, Slide-mode Home navigation, question-only active-Test slides, aligned Review controls, corrected First Wrong navigation, Self Learning filter-only entry until Proceed, the floating current-mode indicator above Zoom, desktop-only CRM with 150% entry and 120%-180% zoom steps, CRM-safe QSVG sizing, a CRM-only Case Study passage fixed to the top 30% of the physical viewport, and a Home control that matches the other CRM rail buttons without double counter-scaling.
Raw LaTeX remains hidden behind dynamically measured skeleton placeholders until MathJax finishes rendering question cards, Solutions, and Explanations.
Initial Attempt badges are synchronized with the selected Question Attempts Max before the HTM is saved.
P/T-based Qtypes are fully data-driven: A/P/T have editable defaults, DOCX codes are detected automatically, custom codes can be added/removed in the GUI, mappings persist in skillnox_qtypes.json, and every Qtype receives a stable distinct button colour in Test Setup and the Self Learning Filter.
QTYPE-column Question Types are also fully data-driven: MCQ/MCQS/TRFL/BLNK/DRWD/MTCH have editable defaults, additional DOCX codes are detected automatically, custom codes can be added/removed, mappings persist in skillnox_question_types.json, and the generated Question Types analytics displays the configured full-form labels while preserving the short codes internally.
The HTML/CSS/JavaScript shell is preserved; document identity, question content, question number, subtopic, level, Question Type/QTYPE metadata, Qtype/P-T metadata, mapped labels, answers, solutions, explanations, and referenced SVG payloads are variable.
"""

from __future__ import annotations

import argparse
import base64
import queue
import shutil
import subprocess
import threading
import webbrowser
import gzip
import hashlib
import html as html_lib
import json
import os
import re
import sys
import traceback
import zipfile
from collections import Counter, OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Union
import xml.etree.ElementTree as ET

APP_NAME = "SkillNox DOCX to HTM"
APP_VERSION = "3.7.0-gui-mode-selector"

OUTPUT_MODE_ORDER: Tuple[str, ...] = ("self", "test", "crm")
OUTPUT_MODE_LABELS: "OrderedDict[str, str]" = OrderedDict((
    ("self", "Self Learning"),
    ("test", "Test"),
    ("crm", "CRM"),
))

# The latest user-approved HTM master shell is compressed and embedded here.
# It includes startup shimmer, current desktop/mobile UI, palette/menu fixes, stable ESVG,
# independent native-sensitivity Slide scrolling, Slide+Palette sticky-left ESVG behavior,
# Slide Home/Test/Review/Self-Learning flow fixes, the current-mode indicator above Zoom,
# rail-based Eye control, dynamic distinct Qtype button colours, desktop-only CRM (150%
# entry; 120%-180% zoom), CRM QSVG viewport fitting, CRM-only 30% Case Study passage
# sizing, and the corrected non-double-scaled Home rail control.
TEMPLATE_GZIP_B64 = (
    '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'
)

REFERENCE_HTML_SHA256 = '51cb8c7bd991f51686d5ea2033b6f03c46fcba93c31e69bfb483e715d30e2f74'

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"

SVG_TOKEN_RE = re.compile(r"#\s*([^#\r\n]+?\.svg)\s*#", re.IGNORECASE)
OPTION_START_RE = re.compile(r"^\s*(\*)?\(([A-Za-z])\)\s*(.*)$")
SOLUTION_MARK_RE = re.compile(r"^\s*Solution\s*:\s*\(([A-Za-z])\)\s*$", re.IGNORECASE)
SOURCE_NUMBER_RE = re.compile(r"^\s*\[\s*([^\]]+)\s*\]\s*")

MAIN_EXPLANATION_HEADINGS = [
    "Explanation:",
    "Concept Idea:",
    "Illustration:",
    "Key Idea:",
    "Data Table:",
    "Units Used:",
    "Theoretical Explanation:",
    "Explanation Step by Step:",
    "Real-Life Application:",
    "Common Mistake:",
    "Final Answer:",
]
MAIN_HEADING_CANON = {h.casefold(): h for h in MAIN_EXPLANATION_HEADINGS}


class GenerationError(RuntimeError):
    """Raised for a source-data or generation problem that must stop output."""


@dataclass
class OptionData:
    letter: str
    lines: List[str]
    starred: bool = False

    @property
    def text(self) -> str:
        return "\n".join(self.lines).strip()


@dataclass
class QuestionData:
    index: int
    source_number: str
    question_lines: List[str]
    options: List[OptionData]
    correct: str
    solution_lines: List[str]
    explanation_raw: str
    qtype: str
    subtopic_raw: str
    subtopic: str
    level: str
    pt: str
    marks: str
    occurrence: str
    qid_raw: str
    qid: str
    se: str
    source_row: List[str] = field(default_factory=list)
    answer_specs: List[str] = field(default_factory=list)
    match_pairs: List[Tuple[str, str]] = field(default_factory=list)
    case_study_id: str = ""
    case_parent_qid: str = ""
    case_parent_subtopic: str = ""
    case_passage_lines: List[str] = field(default_factory=list)
    case_part: str = ""
    case_parts: List[str] = field(default_factory=list)
    case_question_number: int = 0

    @property
    def qkey(self) -> str:
        return f"q{self.index}"

    @property
    def display_number(self) -> str:
        if self.case_study_id and self.case_part and self.case_question_number:
            return f"{self.case_question_number} ({self.case_part})"
        return str(self.index)

    @property
    def palette_number(self) -> Union[int, str]:
        if self.case_study_id and self.case_part and self.case_question_number:
            return f"{self.case_question_number}({self.case_part})"
        return self.index


@dataclass
class GenerationReport:
    exact_reference_mode: bool = False
    exact_reference_match: bool = False
    output_sha256: str = ""
    questions: int = 0
    options: int = 0
    qsvg_references: int = 0
    esvg_references: int = 0
    question_attempt_max: int = 5
    test_attempt_max: int = 10
    test_attempt_rule_count: int = 0
    qtype_count: int = 0
    qtype_codes: List[str] = field(default_factory=list)
    question_type_count: int = 0
    question_type_codes: List[str] = field(default_factory=list)
    enabled_modes: List[str] = field(default_factory=lambda: list(OUTPUT_MODE_ORDER))
    missing_svg_files: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


AttemptValue = Union[int, str]


@dataclass(frozen=True)
class TestAttemptRule:
    from_value: int
    to_value: int
    attempts: AttemptValue

    def as_dict(self) -> Dict[str, object]:
        return {
            "from": self.from_value,
            "to": self.to_value,
            "attempts": self.attempts,
        }


@dataclass(frozen=True)
class AttemptSettings:
    question_attempt_max: int = 5
    test_attempt_max: int = 10
    test_attempt_rules: Tuple[TestAttemptRule, ...] = (
        TestAttemptRule(1, 10, "unlimited"),
        TestAttemptRule(11, 75, 10),
        TestAttemptRule(76, 150, 10),
    )

    def summary(self) -> str:
        return (
            f"Question attempts: {self.question_attempt_max} • "
            f"Default test attempts: {self.test_attempt_max} • "
            f"Test rules: {len(self.test_attempt_rules)}"
        )


DEFAULT_ATTEMPT_SETTINGS = AttemptSettings()


DEFAULT_QTYPE_MAP: "OrderedDict[str, str]" = OrderedDict((
    ("A", "Analytical"),
    ("P", "Problematical"),
    ("T", "Theoretical"),
))
QTYPE_STORE_FILENAME = "skillnox_qtypes.json"


def normalize_qtype_code(value: object) -> str:
    """Normalize a P/T Qtype code while preserving its identity."""
    return str(value or "").strip().upper()


def qtype_store_path() -> Path:
    """Return the persistent Qtype store, preferring the folder beside this script."""
    try:
        script_dir = Path(__file__).resolve().parent
    except Exception:
        script_dir = Path.cwd()
    local_path = script_dir / QTYPE_STORE_FILENAME
    if local_path.exists() or os.access(str(script_dir), os.W_OK):
        return local_path

    appdata = os.environ.get("APPDATA")
    base = Path(appdata).expanduser() / "SkillNox" if appdata else Path.home() / ".skillnox"
    return base / QTYPE_STORE_FILENAME


def load_qtype_store() -> "OrderedDict[str, str]":
    """Load saved Qtypes and merge them over the editable A/P/T defaults."""
    result: "OrderedDict[str, str]" = OrderedDict(DEFAULT_QTYPE_MAP)
    path = qtype_store_path()
    if not path.is_file():
        return result
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        raw_map = payload.get("qtypes", payload) if isinstance(payload, dict) else {}
        if isinstance(raw_map, dict):
            for raw_code, raw_label in raw_map.items():
                code = normalize_qtype_code(raw_code)
                label = str(raw_label or "").strip()
                if code and label:
                    result[code] = label
    except Exception:
        # A damaged optional preferences file must never prevent the generator
        # from starting; the GUI can recreate it from the built-in defaults.
        return OrderedDict(DEFAULT_QTYPE_MAP)
    return result


def save_qtype_store(qtype_map: Dict[str, str]) -> Path:
    """Atomically persist the complete Qtype vocabulary."""
    normalized: "OrderedDict[str, str]" = OrderedDict()
    for code, default_label in DEFAULT_QTYPE_MAP.items():
        label = str(qtype_map.get(code, default_label) or "").strip()
        normalized[code] = label or default_label
    for raw_code, raw_label in qtype_map.items():
        code = normalize_qtype_code(raw_code)
        label = str(raw_label or "").strip()
        if code and code not in normalized and label:
            normalized[code] = label

    path = qtype_store_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": 1,
            "qtypes": normalized,
        }
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        os.replace(temp_path, path)
    except OSError as exc:
        raise GenerationError(f"Could not save Qtype data to {path}: {exc}") from exc
    return path


def qtype_codes_from_questions(questions: Sequence[QuestionData]) -> List[str]:
    """Return normalized P/T codes in first-appearance order."""
    codes: List[str] = []
    seen = set()
    for question in questions:
        code = normalize_qtype_code(question.pt)
        if code and code not in seen:
            seen.add(code)
            codes.append(code)
    return codes


def resolve_qtype_map_for_questions(
    questions: Sequence[QuestionData],
    qtype_map: Optional[Dict[str, str]] = None,
) -> "OrderedDict[str, str]":
    """Validate mappings and return only Qtypes actually used by this DOCX."""
    source = qtype_map if qtype_map is not None else load_qtype_store()
    normalized: Dict[str, str] = {}
    for raw_code, raw_label in source.items():
        code = normalize_qtype_code(raw_code)
        label = str(raw_label or "").strip()
        if code and label:
            normalized[code] = label

    used: "OrderedDict[str, str]" = OrderedDict()
    missing: List[str] = []
    for code in qtype_codes_from_questions(questions):
        label = normalized.get(code, "").strip()
        if not label:
            missing.append(code)
        else:
            used[code] = label
    if missing:
        joined = ", ".join(missing)
        raise GenerationError(
            "Missing Qtype full form for DOCX P/T code(s): " + joined +
            ". Open the GUI and enter the full form, or use --qtype CODE=FULLFORM."
        )
    return used


def detect_qtype_codes_from_docx(docx_path: Path) -> List[str]:
    """Parse the selected DOCX and return the P/T codes it actually uses."""
    tables = read_docx_tables(docx_path)
    table = select_question_table(tables)
    questions = parse_questions(table)
    return qtype_codes_from_questions(questions)


# QTYPE-column Question Type vocabulary (separate from P/T Qtypes).
DEFAULT_QUESTION_TYPE_MAP: "OrderedDict[str, str]" = OrderedDict((
    ("MCQ", "Multiple Choice Question"),
    ("MCQS", "Multiple Choice Questions"),
    ("TRFL", "True or False"),
    ("BLNK", "Fill in the Blanks"),
    ("DRWD", "Drag the Words"),
    ("MTCH", "Match the Following"),
))
QUESTION_TYPE_STORE_FILENAME = "skillnox_question_types.json"


def normalize_question_type_code(value: object) -> str:
    """Normalize a DOCX QTYPE-column code while preserving its short-code identity."""
    return str(value or "").strip().upper()


def question_type_store_path() -> Path:
    """Return the persistent Question Type store beside the script when possible."""
    try:
        script_dir = Path(__file__).resolve().parent
    except Exception:
        script_dir = Path.cwd()
    local_path = script_dir / QUESTION_TYPE_STORE_FILENAME
    if local_path.exists() or os.access(str(script_dir), os.W_OK):
        return local_path

    appdata = os.environ.get("APPDATA")
    base = Path(appdata).expanduser() / "SkillNox" if appdata else Path.home() / ".skillnox"
    return base / QUESTION_TYPE_STORE_FILENAME


def load_question_type_store() -> "OrderedDict[str, str]":
    """Load saved QTYPE-column code/full-form mappings over the built-in defaults."""
    result: "OrderedDict[str, str]" = OrderedDict(DEFAULT_QUESTION_TYPE_MAP)
    path = question_type_store_path()
    if not path.is_file():
        return result
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        raw_map = payload.get("question_types", payload) if isinstance(payload, dict) else {}
        if isinstance(raw_map, dict):
            for raw_code, raw_label in raw_map.items():
                code = normalize_question_type_code(raw_code)
                label = str(raw_label or "").strip()
                if code and label:
                    result[code] = label
    except Exception:
        return OrderedDict(DEFAULT_QUESTION_TYPE_MAP)
    return result


def save_question_type_store(question_type_map: Dict[str, str]) -> Path:
    """Atomically persist the complete QTYPE-column Question Type vocabulary."""
    normalized: "OrderedDict[str, str]" = OrderedDict()
    for code, default_label in DEFAULT_QUESTION_TYPE_MAP.items():
        label = str(question_type_map.get(code, default_label) or "").strip()
        normalized[code] = label or default_label
    for raw_code, raw_label in question_type_map.items():
        code = normalize_question_type_code(raw_code)
        label = str(raw_label or "").strip()
        if code and code not in normalized and label:
            normalized[code] = label

    path = question_type_store_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"version": 1, "question_types": normalized}
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        os.replace(temp_path, path)
    except OSError as exc:
        raise GenerationError(f"Could not save Question Type data to {path}: {exc}") from exc
    return path


def question_type_codes_from_questions(questions: Sequence[QuestionData]) -> List[str]:
    """Return normalized QTYPE-column codes in first-appearance order."""
    codes: List[str] = []
    seen = set()
    for question in questions:
        code = normalize_question_type_code(question.qtype)
        if code and code not in seen:
            seen.add(code)
            codes.append(code)
    return codes


def resolve_question_type_map_for_questions(
    questions: Sequence[QuestionData],
    question_type_map: Optional[Dict[str, str]] = None,
) -> "OrderedDict[str, str]":
    """Validate and return only QTYPE-column mappings used by this DOCX."""
    source = question_type_map if question_type_map is not None else load_question_type_store()
    normalized: Dict[str, str] = {}
    for raw_code, raw_label in source.items():
        code = normalize_question_type_code(raw_code)
        label = str(raw_label or "").strip()
        if code and label:
            normalized[code] = label

    used: "OrderedDict[str, str]" = OrderedDict()
    missing: List[str] = []
    for code in question_type_codes_from_questions(questions):
        label = normalized.get(code, "").strip()
        if not label:
            missing.append(code)
        else:
            used[code] = label
    if missing:
        joined = ", ".join(missing)
        raise GenerationError(
            "Missing Question Type full form for DOCX QTYPE code(s): " + joined +
            ". Open the GUI and enter the full form, or use --question-type CODE=FULLFORM."
        )
    return used


def detect_question_type_codes_from_docx(docx_path: Path) -> List[str]:
    """Parse the selected DOCX and return the QTYPE-column codes it uses."""
    tables = read_docx_tables(docx_path)
    table = select_question_table(tables)
    questions = parse_questions(table)
    return question_type_codes_from_questions(questions)


# ---------------------------------------------------------------------------
# GUI helpers
# ---------------------------------------------------------------------------

def _tk_modules():
    try:
        import tkinter as tk  # type: ignore
        from tkinter import filedialog, messagebox, ttk  # type: ignore
        return tk, ttk, filedialog, messagebox
    except Exception:
        return None, None, None, None


def choose_docx_gui(parent=None) -> Optional[Path]:
    tk, _, filedialog, _ = _tk_modules()
    if tk is None or filedialog is None:
        return None

    owns_root = parent is None
    root = parent or tk.Tk()
    if owns_root:
        root.withdraw()
        root.update_idletasks()
    try:
        selected = filedialog.askopenfilename(
            parent=root,
            title="Select the SkillNox DOCX input file",
            filetypes=[("Word document", "*.docx"), ("All files", "*.*")],
        )
    finally:
        if owns_root:
            root.destroy()
    return Path(selected) if selected else None


def choose_directory_gui(title: str, initial: Path, parent=None) -> Optional[Path]:
    tk, _, filedialog, _ = _tk_modules()
    if tk is None or filedialog is None:
        return None

    owns_root = parent is None
    root = parent or tk.Tk()
    if owns_root:
        root.withdraw()
        root.update_idletasks()
    try:
        selected = filedialog.askdirectory(
            parent=root,
            title=title,
            initialdir=str(initial),
        )
    finally:
        if owns_root:
            root.destroy()
    return Path(selected) if selected else None


def show_message(kind: str, title: str, message: str, parent=None) -> None:
    tk, _, _, messagebox = _tk_modules()
    if tk is None or messagebox is None:
        stream = sys.stderr if kind == "error" else sys.stdout
        print(f"{title}: {message}", file=stream)
        return

    owns_root = parent is None
    root = parent or tk.Tk()
    if owns_root:
        root.withdraw()
        root.update_idletasks()
    try:
        if kind == "error":
            messagebox.showerror(title, message, parent=root)
        elif kind == "warning":
            messagebox.showwarning(title, message, parent=root)
        else:
            messagebox.showinfo(title, message, parent=root)
    finally:
        if owns_root:
            root.destroy()


def _windows_chrome_from_registry() -> Optional[Path]:
    if not sys.platform.startswith("win"):
        return None
    try:
        import winreg  # type: ignore
    except Exception:
        return None

    key_names = (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
    )
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for key_name in key_names:
            try:
                with winreg.OpenKey(hive, key_name) as key:
                    value, _ = winreg.QueryValueEx(key, None)
                candidate = Path(str(value).strip('"')).expanduser()
                if candidate.is_file():
                    return candidate.resolve()
            except OSError:
                continue
            except Exception:
                continue
    return None


def find_chrome_executable() -> Optional[Path]:
    """Find Google Chrome without requiring a third-party package."""
    registry_path = _windows_chrome_from_registry()
    if registry_path is not None:
        return registry_path

    candidates: List[Path] = []
    if sys.platform.startswith("win"):
        for env_name in ("LOCALAPPDATA", "PROGRAMFILES", "PROGRAMFILES(X86)", "PROGRAMW6432"):
            base = os.environ.get(env_name)
            if base:
                candidates.append(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe")
    elif sys.platform == "darwin":
        candidates.extend([
            Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
            Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ])

    for candidate in candidates:
        try:
            if candidate.is_file():
                return candidate.resolve()
        except OSError:
            continue

    executable_names = (
        "chrome.exe",
        "chrome",
        "google-chrome",
        "google-chrome-stable",
        "chromium",
        "chromium-browser",
    )
    for name in executable_names:
        found = shutil.which(name)
        if found:
            return Path(found).resolve()
    return None


def open_html_in_chrome(output_path: Path) -> Tuple[bool, str]:
    """Open the generated file in Chrome; fall back to the default browser."""
    output_path = output_path.expanduser().resolve()
    url = output_path.as_uri()
    chrome = find_chrome_executable()

    try:
        if chrome is not None:
            kwargs = {
                "stdout": subprocess.DEVNULL,
                "stderr": subprocess.DEVNULL,
            }
            if sys.platform.startswith("win") and hasattr(subprocess, "CREATE_NO_WINDOW"):
                kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW  # type: ignore[index]
            subprocess.Popen([str(chrome), "--new-tab", url], **kwargs)
            return True, f"Opened automatically in Google Chrome:\n{chrome}"

        # macOS can resolve the application even when the executable path is not conventional.
        if sys.platform == "darwin" and shutil.which("open"):
            subprocess.Popen(
                ["open", "-a", "Google Chrome", url],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True, "Opened automatically in Google Chrome."
    except Exception as exc:
        chrome_error = str(exc)
    else:
        chrome_error = "Google Chrome was not found."

    try:
        opened = bool(webbrowser.open_new_tab(url))
    except Exception as exc:
        return False, f"Could not open the generated HTM in Chrome or the default browser: {exc}"

    if opened:
        return False, f"{chrome_error} The HTM was opened in the system default browser instead."
    return False, f"{chrome_error} Open the file manually:\n{output_path}"


def _set_windows_dpi_awareness() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class SkillNoxGeneratorGUI:
    """Modern Tkinter front end for the unchanged SkillNox HTML generator."""

    COLORS = {
        "bg": "#07111F",
        "panel": "#0A1726",
        "card": "#0F2032",
        "card_alt": "#0B1A29",
        "field": "#081624",
        "border": "#213D56",
        "border_soft": "#173047",
        "text": "#ECF5FF",
        "muted": "#8FA7BB",
        "accent": "#2DD4BF",
        "accent_hover": "#5EEAD4",
        "blue": "#38BDF8",
        "violet": "#8B5CF6",
        "success": "#34D399",
        "warning": "#FBBF24",
        "danger": "#FB7185",
        "button_text": "#04151B",
    }

    DEFAULT_RULES = (
        (1, 10, "Unlimited"),
        (11, 75, "10"),
        (76, 150, "10"),
    )

    def __init__(self, root, args) -> None:
        self.root = root
        self.args = args
        self.result_queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.selected_docx: Optional[Path] = None
        self.images_q: Optional[Path] = None
        self.images_e: Optional[Path] = None
        self.output_path: Optional[Path] = None
        self.selected_question_count = 0
        self.exit_code = 0
        self.busy = False
        self.rule_rows: List[Dict[str, object]] = []
        self.qtype_rows: List[Dict[str, object]] = []
        self.detected_qtype_codes: List[str] = []
        self.qtype_store: "OrderedDict[str, str]" = load_qtype_store()
        for code, label in (getattr(args, "qtype", None) or []):
            self.qtype_store[normalize_qtype_code(code)] = str(label).strip()

        self.question_type_rows: List[Dict[str, object]] = []
        self.detected_question_type_codes: List[str] = []
        self.question_type_store: "OrderedDict[str, str]" = load_question_type_store()
        for code, label in (getattr(args, "question_type", None) or []):
            self.question_type_store[normalize_question_type_code(code)] = str(label).strip()

        self.config_widgets: List[object] = []
        self._settings_refresh_job = None
        self._qtype_refresh_job = None
        self._qtype_refreshing = False
        self._question_type_refresh_job = None
        self._question_type_refreshing = False

        tk, ttk, _, _ = _tk_modules()
        if tk is None or ttk is None:
            raise RuntimeError("Tkinter is unavailable")
        self.tk = tk
        self.ttk = ttk

        self.docx_var = tk.StringVar(value="No DOCX selected")
        self.resources_var = tk.StringVar(value="imagesQ and imagesE will be detected beside the DOCX")
        self.output_var = tk.StringVar(value="Output: <DOCX name>_generated.htm")
        self.status_var = tk.StringVar(value="Select a DOCX, choose HTM mode(s), review settings, then generate")
        self.settings_feedback_var = tk.StringVar(value="Default attempt rules are ready")
        self.qtype_feedback_var = tk.StringVar(value="Qtype vocabulary is ready")
        self.question_type_feedback_var = tk.StringVar(value="Question Type vocabulary is ready")
        self.question_count_var = tk.StringVar(value="Total scored questions: —")
        self.mode_feedback_var = tk.StringVar(value="Select at least one HTM mode")
        self.output_mode_vars: "OrderedDict[str, object]" = OrderedDict(
            (mode, tk.BooleanVar(value=False)) for mode in OUTPUT_MODE_ORDER
        )
        self.output_mode_buttons: Dict[str, object] = {}
        self.question_attempt_max_var = tk.StringVar(value="5")
        self.test_attempt_max_var = tk.StringVar(value="10")

        self._configure_window()
        self._build_ui()
        self._populate_qtype_rows()
        self._populate_question_type_rows()
        self._reset_default_rules(initial=True)
        self.root.after(220, self._initial_select)
        self.root.after(100, self._poll_results)

    # --------------------------- window + styling ---------------------------

    def _configure_window(self) -> None:
        c = self.COLORS
        self.root.title(f"{APP_NAME} {APP_VERSION}")
        self.root.geometry("1100x900")
        self.root.minsize(900, 680)
        self.root.configure(bg=c["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        style = self.ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Smart.Horizontal.TProgressbar",
            troughcolor=c["field"],
            background=c["accent"],
            bordercolor=c["field"],
            lightcolor=c["accent"],
            darkcolor=c["accent"],
            thickness=8,
        )
        style.configure(
            "Smart.Vertical.TScrollbar",
            troughcolor=c["bg"],
            background=c["border"],
            bordercolor=c["bg"],
            arrowcolor=c["muted"],
        )

    def _font(self, size: int, weight: str = "normal"):
        return ("Segoe UI", size, weight)

    def _button(
        self,
        parent,
        text: str,
        command,
        *,
        kind: str = "secondary",
        width: Optional[int] = None,
        padx: int = 18,
        pady: int = 10,
    ):
        c = self.COLORS
        palette = {
            "primary": (c["accent"], c["button_text"], c["accent_hover"]),
            "secondary": (c["card_alt"], c["text"], c["border"]),
            "danger": ("#3A1724", "#FFDCE6", "#572034"),
            "ghost": (c["panel"], c["muted"], c["card_alt"]),
        }
        bg, fg, active = palette.get(kind, palette["secondary"])
        button = self.tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active,
            activeforeground=fg,
            disabledforeground="#5D7185",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=c["border"],
            highlightcolor=c["blue"],
            cursor="hand2",
            font=self._font(10, "bold"),
            padx=padx,
            pady=pady,
            width=width,
        )
        return button

    def _entry(self, parent, variable, *, width: Optional[int] = None, justify: str = "left"):
        c = self.COLORS
        entry = self.tk.Entry(
            parent,
            textvariable=variable,
            width=width,
            justify=justify,
            bg=c["field"],
            fg=c["text"],
            insertbackground=c["text"],
            selectbackground=c["violet"],
            selectforeground="#FFFFFF",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=c["border"],
            highlightcolor=c["blue"],
            font=self._font(10, "bold"),
        )
        return entry

    def _make_card(self, parent, title: str, subtitle: str = "", accent: Optional[str] = None):
        c = self.COLORS
        card = self.tk.Frame(
            parent,
            bg=c["card"],
            highlightthickness=1,
            highlightbackground=c["border"],
            bd=0,
        )
        header = self.tk.Frame(card, bg=c["card"])
        header.pack(fill="x", padx=20, pady=(18, 10))
        if accent:
            self.tk.Frame(header, bg=accent, width=4, height=40).pack(side="left", padx=(0, 12))
        text_box = self.tk.Frame(header, bg=c["card"])
        text_box.pack(side="left", fill="x", expand=True)
        self.tk.Label(
            text_box,
            text=title,
            bg=c["card"],
            fg=c["text"],
            font=self._font(13, "bold"),
        ).pack(anchor="w")
        if subtitle:
            self.tk.Label(
                text_box,
                text=subtitle,
                bg=c["card"],
                fg=c["muted"],
                font=self._font(9),
                justify="left",
                wraplength=820,
            ).pack(anchor="w", pady=(3, 0))
        body = self.tk.Frame(card, bg=c["card"])
        body.pack(fill="both", expand=True, padx=20, pady=(0, 18))
        return card, header, body

    def _build_ui(self) -> None:
        c = self.COLORS

        shell = self.tk.Frame(self.root, bg=c["bg"])
        shell.pack(fill="both", expand=True)

        # Header area
        header = self.tk.Frame(shell, bg=c["panel"], highlightthickness=1, highlightbackground=c["border_soft"])
        header.pack(fill="x")
        header_inner = self.tk.Frame(header, bg=c["panel"])
        header_inner.pack(fill="x", padx=28, pady=18)

        logo = self.tk.Label(
            header_inner,
            text="SN",
            bg=c["violet"],
            fg="#FFFFFF",
            width=3,
            height=1,
            font=self._font(16, "bold"),
        )
        logo.pack(side="left", padx=(0, 14))

        title_box = self.tk.Frame(header_inner, bg=c["panel"])
        title_box.pack(side="left", fill="x", expand=True)
        self.tk.Label(
            title_box,
            text="SkillNox DOCX → HTM Studio",
            bg=c["panel"],
            fg=c["text"],
            font=self._font(22, "bold"),
        ).pack(anchor="w")
        self.tk.Label(
            title_box,
            text="Generate the validated quiz shell, show the DOCX question total, choose Self Learning/Test/CRM modes, configure attempts, and open the result in Chrome.",
            bg=c["panel"],
            fg=c["muted"],
            font=self._font(10),
        ).pack(anchor="w", pady=(3, 0))

        self.tk.Label(
            header_inner,
            text=f"v{APP_VERSION}",
            bg=c["card_alt"],
            fg=c["accent"],
            font=self._font(9, "bold"),
            padx=12,
            pady=7,
        ).pack(side="right")

        feature_bar = self.tk.Frame(header, bg=c["panel"])
        feature_bar.pack(fill="x", padx=28, pady=(0, 18))
        features = (
            ("DOCX variable data", c["blue"]),
            ("Dynamic Qtypes", c["accent"]),
            ("Question Type codes", c["blue"]),
            ("QSVG + ESVG inlining", c["violet"]),
            ("No Base64 in HTM", c["success"]),
            ("Auto-open Chrome", c["warning"]),
        )
        for text, color in features:
            self.tk.Label(
                feature_bar,
                text="●  " + text,
                bg=c["card_alt"],
                fg=color,
                font=self._font(9, "bold"),
                padx=12,
                pady=6,
            ).pack(side="left", padx=(0, 8))

        # Scrollable content area
        content_shell = self.tk.Frame(shell, bg=c["bg"])
        content_shell.pack(fill="both", expand=True)
        self.canvas = self.tk.Canvas(content_shell, bg=c["bg"], highlightthickness=0, bd=0)
        scrollbar = self.ttk.Scrollbar(
            content_shell,
            orient="vertical",
            command=self.canvas.yview,
            style="Smart.Vertical.TScrollbar",
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.content = self.tk.Frame(self.canvas, bg=c["bg"])
        self.content_window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.content.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(self.content_window, width=e.width),
        )

        def on_mousewheel(event):
            if sys.platform == "darwin":
                delta = -1 if event.delta > 0 else 1
            else:
                delta = int(-1 * (event.delta / 120)) if event.delta else 0
            if delta:
                self.canvas.yview_scroll(delta, "units")

        self.canvas.bind_all("<MouseWheel>", on_mousewheel)
        self.canvas.bind_all("<Button-4>", lambda _e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind_all("<Button-5>", lambda _e: self.canvas.yview_scroll(1, "units"))

        page = self.tk.Frame(self.content, bg=c["bg"])
        page.pack(fill="both", expand=True, padx=28, pady=24)

        # Input card
        file_card, _, file_body = self._make_card(
            page,
            "1. Source document",
            "Select the SkillNox question-bank DOCX. imagesQ and imagesE are detected beside it.",
            c["blue"],
        )
        file_card.pack(fill="x")

        file_row = self.tk.Frame(file_body, bg=c["card"])
        file_row.pack(fill="x")
        path_box = self.tk.Frame(file_row, bg=c["field"], highlightthickness=1, highlightbackground=c["border"])
        path_box.pack(side="left", fill="x", expand=True)
        self.tk.Label(
            path_box,
            textvariable=self.docx_var,
            bg=c["field"],
            fg=c["text"],
            font=self._font(10, "bold"),
            anchor="w",
            justify="left",
            padx=14,
            pady=12,
            wraplength=710,
        ).pack(fill="x")
        self.select_button = self._button(
            file_row,
            "Choose DOCX",
            lambda: self.select_docx(auto_start=False),
            kind="secondary",
            padx=18,
            pady=11,
        )
        self.select_button.pack(side="right", padx=(14, 0))

        info_grid = self.tk.Frame(file_body, bg=c["card"])
        info_grid.pack(fill="x", pady=(12, 0))
        info_grid.columnconfigure(0, weight=1)
        info_grid.columnconfigure(1, weight=1)
        resource_box = self.tk.Frame(info_grid, bg=c["card_alt"], highlightthickness=1, highlightbackground=c["border_soft"])
        resource_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        output_box = self.tk.Frame(info_grid, bg=c["card_alt"], highlightthickness=1, highlightbackground=c["border_soft"])
        output_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.tk.Label(resource_box, text="Detected resources", bg=c["card_alt"], fg=c["muted"], font=self._font(8, "bold"), padx=12, pady=6).pack(anchor="w")
        self.tk.Label(resource_box, textvariable=self.resources_var, bg=c["card_alt"], fg=c["text"], font=self._font(9), justify="left", anchor="w", padx=12, pady=0, wraplength=420).pack(fill="x", pady=(0, 10))
        self.tk.Label(output_box, text="Output file", bg=c["card_alt"], fg=c["muted"], font=self._font(8, "bold"), padx=12, pady=6).pack(anchor="w")
        self.tk.Label(output_box, textvariable=self.output_var, bg=c["card_alt"], fg=c["text"], font=self._font(9), justify="left", anchor="w", padx=12, pady=0, wraplength=420).pack(fill="x", pady=(0, 10))

        # Qtype management card
        qtype_card, qtype_header, qtype_body = self._make_card(
            page,
            "2. Qtype management",
            "P/T codes are detected from the DOCX. Add your own codes with + Add Qtype; complete mappings are saved automatically for future files.",
            c["blue"],
        )
        qtype_card.pack(fill="x", pady=(18, 0))
        self.add_qtype_button = self._button(
            qtype_header,
            "+ Add Qtype",
            self._add_qtype_row,
            kind="secondary",
            padx=13,
            pady=7,
        )
        self.add_qtype_button.pack(side="right")
        self.config_widgets.append(self.add_qtype_button)

        qtype_table = self.tk.Frame(qtype_body, bg=c["card_alt"], highlightthickness=1, highlightbackground=c["border"])
        qtype_table.pack(fill="x")
        qtype_header_row = self.tk.Frame(qtype_table, bg="#11283A")
        qtype_header_row.pack(fill="x")
        qtype_headers = (("Code", 12), ("Full form", 44), ("Status", 18), ("", 9))
        for index, (label, width) in enumerate(qtype_headers):
            cell = self.tk.Label(
                qtype_header_row,
                text=label,
                width=width,
                bg="#11283A",
                fg=c["muted"],
                font=self._font(8, "bold"),
                pady=8,
            )
            cell.grid(row=0, column=index, sticky="ew")
        qtype_header_row.columnconfigure(1, weight=1)

        self.qtype_container = self.tk.Frame(qtype_table, bg=c["card_alt"])
        self.qtype_container.pack(fill="x")

        qtype_hint_row = self.tk.Frame(qtype_body, bg=c["card"])
        qtype_hint_row.pack(fill="x", pady=(10, 0))
        self.tk.Label(
            qtype_hint_row,
            text="A/P/T are built-in and cannot be removed, but their full forms are editable. A custom Qtype used by the selected DOCX cannot be removed.",
            bg=c["card"],
            fg=c["muted"],
            font=self._font(8),
            justify="left",
            wraplength=720,
        ).pack(side="left", fill="x", expand=True)
        self.qtype_feedback_label = self.tk.Label(
            qtype_hint_row,
            textvariable=self.qtype_feedback_var,
            bg=c["card_alt"],
            fg=c["success"],
            font=self._font(8, "bold"),
            padx=10,
            pady=5,
        )
        self.qtype_feedback_label.pack(side="right", padx=(12, 0))

        # QTYPE-column Question Type management card
        question_type_card, question_type_header, question_type_body = self._make_card(
            page,
            "3. Question Type management",
            "QTYPE-column short codes such as MCQ, TRFL, BLNK, DRWD, and MTCH are detected from the DOCX. Add custom codes and edit their full-form labels; mappings are saved automatically.",
            c["accent"],
        )
        question_type_card.pack(fill="x", pady=(18, 0))
        self.add_question_type_button = self._button(
            question_type_header,
            "+ Add Question Type",
            self._add_question_type_row,
            kind="secondary",
            padx=13,
            pady=7,
        )
        self.add_question_type_button.pack(side="right")
        self.config_widgets.append(self.add_question_type_button)

        question_type_table = self.tk.Frame(question_type_body, bg=c["card_alt"], highlightthickness=1, highlightbackground=c["border"])
        question_type_table.pack(fill="x")
        question_type_header_row = self.tk.Frame(question_type_table, bg="#11283A")
        question_type_header_row.pack(fill="x")
        question_type_headers = (("Code", 12), ("Full form", 44), ("Status", 18), ("", 9))
        for index, (label, width) in enumerate(question_type_headers):
            cell = self.tk.Label(
                question_type_header_row,
                text=label,
                width=width,
                bg="#11283A",
                fg=c["muted"],
                font=self._font(8, "bold"),
                pady=8,
            )
            cell.grid(row=0, column=index, sticky="ew")
        question_type_header_row.columnconfigure(1, weight=1)

        self.question_type_container = self.tk.Frame(question_type_table, bg=c["card_alt"])
        self.question_type_container.pack(fill="x")

        question_type_hint_row = self.tk.Frame(question_type_body, bg=c["card"])
        question_type_hint_row.pack(fill="x", pady=(10, 0))
        self.tk.Label(
            question_type_hint_row,
            text="MCQ/MCQS/TRFL/BLNK/DRWD/MTCH are built-in and cannot be removed, but their full forms are editable. A custom Question Type used by the selected DOCX cannot be removed.",
            bg=c["card"],
            fg=c["muted"],
            font=self._font(8),
            justify="left",
            wraplength=720,
        ).pack(side="left", fill="x", expand=True)
        self.question_type_feedback_label = self.tk.Label(
            question_type_hint_row,
            textvariable=self.question_type_feedback_var,
            bg=c["card_alt"],
            fg=c["success"],
            font=self._font(8, "bold"),
            padx=10,
            pady=5,
        )
        self.question_type_feedback_label.pack(side="right", padx=(12, 0))

        # Attempts configuration card
        config_card, config_header, config_body = self._make_card(
            page,
            "4. Attempts configuration",
            "The detected scored-question total is shown first. Attempt values are written into window.skillnoxAttemptConfig.",
            c["violet"],
        )
        config_card.pack(fill="x", pady=(18, 0))
        reset_button = self._button(config_header, "Reset defaults", self._reset_default_rules, kind="ghost", padx=12, pady=7)
        reset_button.pack(side="right")
        self.config_widgets.append(reset_button)

        question_count_box = self.tk.Frame(
            config_body,
            bg="#10293A",
            highlightthickness=1,
            highlightbackground=c["blue"],
        )
        question_count_box.pack(fill="x", pady=(0, 14))
        self.tk.Label(
            question_count_box,
            text="DOCX QUESTION COUNT",
            bg="#10293A",
            fg=c["muted"],
            font=self._font(8, "bold"),
            padx=14,
            pady=5,
        ).pack(anchor="w")
        self.tk.Label(
            question_count_box,
            textvariable=self.question_count_var,
            bg="#10293A",
            fg=c["text"],
            font=self._font(15, "bold"),
            padx=14,
            pady=4,
        ).pack(anchor="w", pady=(0, 8))

        limits_grid = self.tk.Frame(config_body, bg=c["card"])
        limits_grid.pack(fill="x")
        limits_grid.columnconfigure(0, weight=1)
        limits_grid.columnconfigure(1, weight=1)

        qmax_row = self.tk.Frame(limits_grid, bg=c["card_alt"], highlightthickness=1, highlightbackground=c["border_soft"])
        qmax_row.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        qmax_copy = self.tk.Frame(qmax_row, bg=c["card_alt"])
        qmax_copy.pack(side="left", fill="x", expand=True, padx=16, pady=14)
        self.tk.Label(qmax_copy, text="Question Attempts Max", bg=c["card_alt"], fg=c["text"], font=self._font(11, "bold")).pack(anchor="w")
        self.tk.Label(
            qmax_copy,
            text="Attempts allowed for each individual question (1–99).",
            bg=c["card_alt"],
            fg=c["muted"],
            font=self._font(9),
        ).pack(anchor="w", pady=(3, 0))
        qmax_entry = self._entry(qmax_row, self.question_attempt_max_var, width=7, justify="center")
        qmax_entry.pack(side="right", padx=16, ipady=9)
        self.config_widgets.append(qmax_entry)
        self.question_attempt_max_var.trace_add("write", lambda *_: self._schedule_settings_refresh())

        tmax_row = self.tk.Frame(limits_grid, bg=c["card_alt"], highlightthickness=1, highlightbackground=c["border_soft"])
        tmax_row.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        tmax_copy = self.tk.Frame(tmax_row, bg=c["card_alt"])
        tmax_copy.pack(side="left", fill="x", expand=True, padx=16, pady=14)
        self.tk.Label(tmax_copy, text="Default Test Attempts Max", bg=c["card_alt"], fg=c["text"], font=self._font(11, "bold")).pack(anchor="w")
        self.tk.Label(
            tmax_copy,
            text="Fallback test re-attempt limit when no range rule matches (1–999).",
            bg=c["card_alt"],
            fg=c["muted"],
            font=self._font(9),
            wraplength=330,
            justify="left",
        ).pack(anchor="w", pady=(3, 0))
        tmax_entry = self._entry(tmax_row, self.test_attempt_max_var, width=7, justify="center")
        tmax_entry.pack(side="right", padx=16, ipady=9)
        self.config_widgets.append(tmax_entry)
        self.test_attempt_max_var.trace_add("write", lambda *_: self._schedule_settings_refresh())

        rules_head = self.tk.Frame(config_body, bg=c["card"])
        rules_head.pack(fill="x", pady=(18, 8))
        self.tk.Label(rules_head, text="Test Attempt Rules", bg=c["card"], fg=c["text"], font=self._font(11, "bold")).pack(side="left")
        self.tk.Label(
            rules_head,
            text="Matched against the selected test question-count range",
            bg=c["card"],
            fg=c["muted"],
            font=self._font(9),
        ).pack(side="left", padx=(10, 0))
        add_button = self._button(rules_head, "+ Add Rule", self._add_rule, kind="secondary", padx=13, pady=7)
        add_button.pack(side="right")
        self.add_rule_button = add_button
        self.config_widgets.append(add_button)

        table = self.tk.Frame(config_body, bg=c["card_alt"], highlightthickness=1, highlightbackground=c["border"])
        table.pack(fill="x")
        header_row = self.tk.Frame(table, bg="#11283A")
        header_row.pack(fill="x")
        headers = (("#", 5), ("From", 15), ("To", 15), ("Attempts", 24), ("", 6))
        for index, (label, width) in enumerate(headers):
            cell = self.tk.Label(
                header_row,
                text=label,
                width=width,
                bg="#11283A",
                fg=c["muted"],
                font=self._font(8, "bold"),
                pady=8,
            )
            cell.grid(row=0, column=index, sticky="ew")
        header_row.columnconfigure(3, weight=1)

        self.rules_container = self.tk.Frame(table, bg=c["card_alt"])
        self.rules_container.pack(fill="x")

        hint_row = self.tk.Frame(config_body, bg=c["card"])
        hint_row.pack(fill="x", pady=(10, 0))
        self.tk.Label(
            hint_row,
            text="Attempts accepts Unlimited, 0 (unlimited), or a positive integer. New rows continue from the previous maximum To value.",
            bg=c["card"],
            fg=c["muted"],
            font=self._font(8),
            justify="left",
            wraplength=760,
        ).pack(side="left", fill="x", expand=True)
        self.settings_feedback_label = self.tk.Label(
            hint_row,
            textvariable=self.settings_feedback_var,
            bg=c["card_alt"],
            fg=c["success"],
            font=self._font(8, "bold"),
            padx=10,
            pady=5,
        )
        self.settings_feedback_label.pack(side="right", padx=(12, 0))

        # Output-mode selection card
        mode_card, _, mode_body = self._make_card(
            page,
            "5. HTM modes",
            "Select one or more modes. All three selected keeps the full reference behavior. If only one is selected, the generated HTM exposes only that mode and opens directly in it.",
            c["warning"],
        )
        mode_card.pack(fill="x", pady=(18, 0))

        mode_grid = self.tk.Frame(mode_body, bg=c["card"])
        mode_grid.pack(fill="x")
        for column in range(3):
            mode_grid.columnconfigure(column, weight=1)
        mode_specs = (
            ("self", "Self Learning", c["success"]),
            ("test", "Test", c["blue"]),
            ("crm", "CRM", c["warning"]),
        )
        for column, (mode, label, accent) in enumerate(mode_specs):
            button = self.tk.Button(
                mode_grid,
                text=label,
                command=lambda key=mode: self._toggle_output_mode(key),
                bg=c["card_alt"],
                fg=c["text"],
                activebackground=accent,
                activeforeground=c["bg"],
                relief="flat",
                bd=0,
                highlightthickness=1,
                highlightbackground=c["border"],
                highlightcolor=accent,
                cursor="hand2",
                font=self._font(11, "bold"),
                padx=16,
                pady=14,
            )
            button.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 6, 0 if column == 2 else 6))
            self.output_mode_buttons[mode] = button
            self.config_widgets.append(button)

        self.mode_feedback_label = self.tk.Label(
            mode_body,
            textvariable=self.mode_feedback_var,
            bg=c["card_alt"],
            fg=c["warning"],
            font=self._font(9, "bold"),
            padx=12,
            pady=7,
        )
        self.mode_feedback_label.pack(anchor="w", pady=(12, 0))
        self._refresh_output_mode_ui()

        # Processing card
        process_card, _, process_body = self._make_card(
            page,
            "6. Generate and open",
            "The finished HTM is saved beside the DOCX and opened automatically in Google Chrome.",
            c["accent"],
        )
        process_card.pack(fill="x", pady=(18, 0))
        status_row = self.tk.Frame(process_body, bg=c["card"])
        status_row.pack(fill="x")
        self.status_dot = self.tk.Label(status_row, text="●", bg=c["card"], fg=c["blue"], font=self._font(12, "bold"))
        self.status_dot.pack(side="left", padx=(0, 8))
        self.status_label = self.tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=c["card"],
            fg=c["text"],
            font=self._font(10, "bold"),
            anchor="w",
        )
        self.status_label.pack(side="left", fill="x", expand=True)
        self.progress = self.ttk.Progressbar(process_body, mode="indeterminate", style="Smart.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(12, 0))

        # Fixed bottom action bar
        action_bar = self.tk.Frame(shell, bg=c["panel"], highlightthickness=1, highlightbackground=c["border_soft"])
        action_bar.pack(fill="x", side="bottom")
        action_inner = self.tk.Frame(action_bar, bg=c["panel"])
        action_inner.pack(fill="x", padx=28, pady=14)
        self.tk.Label(
            action_inner,
            text="Validated no-Base64 HTM shell • Chrome opens after successful generation",
            bg=c["panel"],
            fg=c["muted"],
            font=self._font(9),
        ).pack(side="left")
        self.exit_button = self._button(action_inner, "Close", self._on_close, kind="ghost", padx=18, pady=10)
        self.exit_button.pack(side="right")
        self.generate_button = self._button(
            action_inner,
            "Generate HTM & Open Chrome",
            self.start_generation,
            kind="primary",
            padx=22,
            pady=10,
        )
        self.generate_button.pack(side="right", padx=(0, 10))
        self.generate_button.configure(state="disabled")

    # --------------------------- Qtype management UI ---------------------------

    def _populate_qtype_rows(self) -> None:
        for row in list(self.qtype_rows):
            frame = row.get("frame")
            if frame is not None:
                try:
                    frame.destroy()
                except Exception:
                    pass
        self.qtype_rows.clear()

        for code, full_form in self.qtype_store.items():
            origin = "builtin" if code in DEFAULT_QTYPE_MAP else "saved"
            self._add_qtype_row(code=code, full_form=full_form, origin=origin, refresh=False)
        self._refresh_qtype_state(save_if_valid=False)

    def _add_qtype_row(
        self,
        code: str = "",
        full_form: str = "",
        origin: str = "manual",
        refresh: bool = True,
    ) -> None:
        c = self.COLORS
        row_frame = self.tk.Frame(self.qtype_container, bg=c["card_alt"])
        row_frame.pack(fill="x", padx=8, pady=5)

        code_var = self.tk.StringVar(value=normalize_qtype_code(code) if code else "")
        full_form_var = self.tk.StringVar(value=str(full_form or ""))
        code_entry = self._entry(row_frame, code_var, width=10, justify="center")
        code_entry.grid(row=0, column=0, sticky="ew", padx=(0, 4), ipady=7)
        full_form_entry = self._entry(row_frame, full_form_var, width=38, justify="left")
        full_form_entry.grid(row=0, column=1, sticky="ew", padx=4, ipady=7)

        status_label = self.tk.Label(
            row_frame,
            text="New",
            width=18,
            bg=c["card_alt"],
            fg=c["muted"],
            font=self._font(8, "bold"),
        )
        status_label.grid(row=0, column=2, sticky="ew", padx=4)

        row: Dict[str, object] = {
            "frame": row_frame,
            "code_var": code_var,
            "full_form_var": full_form_var,
            "code_entry": code_entry,
            "full_form_entry": full_form_entry,
            "status_label": status_label,
            "origin": origin,
        }

        if origin == "builtin":
            self.tk.Label(
                row_frame,
                text="Built-in",
                width=8,
                bg=c["card_alt"],
                fg=c["muted"],
                font=self._font(8, "bold"),
            ).grid(row=0, column=3, sticky="e", padx=(5, 0))
        else:
            remove_button = self._button(
                row_frame,
                "Remove",
                lambda r=row: self._remove_qtype_row(r),
                kind="danger",
                width=7,
                padx=8,
                pady=7,
            )
            remove_button.grid(row=0, column=3, sticky="e", padx=(5, 0))
            row["remove_button"] = remove_button
            self.config_widgets.append(remove_button)

        row_frame.columnconfigure(1, weight=1)
        self.qtype_rows.append(row)
        self.config_widgets.extend([code_entry, full_form_entry])

        code_var.trace_add("write", lambda *_: self._schedule_qtype_refresh())
        full_form_var.trace_add("write", lambda *_: self._schedule_qtype_refresh())
        code_entry.bind("<FocusOut>", lambda _e, r=row: self._normalize_qtype_row_code(r))

        self._refresh_qtype_row_badges()
        if refresh:
            self._schedule_qtype_refresh()

    def _normalize_qtype_row_code(self, row: Dict[str, object]) -> None:
        if self._qtype_refreshing:
            return
        code_var = row.get("code_var")
        if code_var is None:
            return
        try:
            normalized = normalize_qtype_code(code_var.get())
            if normalized != str(code_var.get()):
                code_var.set(normalized)
        except Exception:
            pass

    def _qtype_row_code(self, row: Dict[str, object]) -> str:
        try:
            return normalize_qtype_code(row["code_var"].get())
        except Exception:
            return ""

    def _find_qtype_row(self, code: str) -> Optional[Dict[str, object]]:
        wanted = normalize_qtype_code(code)
        for row in self.qtype_rows:
            if self._qtype_row_code(row) == wanted:
                return row
        return None

    def _remove_qtype_row(self, row: Dict[str, object]) -> None:
        if row not in self.qtype_rows:
            return
        code = self._qtype_row_code(row)
        if code and code in self.detected_qtype_codes:
            show_message(
                "error",
                "Qtype is in use",
                f"Cannot remove Qtype {code} because it is used by the selected DOCX.",
                parent=self.root,
            )
            return
        if str(row.get("origin", "")) == "builtin":
            return

        self.qtype_rows.remove(row)
        frame = row.get("frame")
        if frame is not None:
            try:
                frame.destroy()
            except Exception:
                pass
        self._refresh_qtype_state(save_if_valid=True)

    def _sync_detected_qtypes(self, codes: Sequence[str]) -> None:
        normalized: List[str] = []
        seen = set()
        for value in codes:
            code = normalize_qtype_code(value)
            if code and code not in seen:
                seen.add(code)
                normalized.append(code)
        self.detected_qtype_codes = normalized

        # Drop only temporary DOCX-only blank rows that are no longer needed.
        for row in list(self.qtype_rows):
            if str(row.get("origin", "")) != "detected":
                continue
            code = self._qtype_row_code(row)
            if code in self.detected_qtype_codes or code in self.qtype_store:
                continue
            self.qtype_rows.remove(row)
            frame = row.get("frame")
            if frame is not None:
                try:
                    frame.destroy()
                except Exception:
                    pass

        for code in self.detected_qtype_codes:
            if self._find_qtype_row(code) is None:
                self._add_qtype_row(code=code, full_form="", origin="detected", refresh=False)

        self._refresh_qtype_row_badges()
        self._refresh_qtype_state(save_if_valid=False)

    def _collect_qtype_map(self, show_errors: bool = False) -> Optional["OrderedDict[str, str]"]:
        errors: List[str] = []
        mapping: "OrderedDict[str, str]" = OrderedDict()

        for index, row in enumerate(self.qtype_rows, start=1):
            code = self._qtype_row_code(row)
            try:
                full_form = str(row["full_form_var"].get()).strip()
            except Exception:
                full_form = ""

            if not code and not full_form:
                errors.append(f"Qtype row {index}: enter a code and full form, or remove the row.")
                continue
            if not code:
                errors.append(f"Qtype row {index}: Code cannot be blank.")
                continue
            if not full_form:
                errors.append(f"Qtype {code}: Full form cannot be blank.")
                continue
            if code in mapping:
                errors.append(f"Qtype {code} already exists. Duplicate codes are not allowed.")
                continue
            mapping[code] = full_form

        for code in self.detected_qtype_codes:
            if not mapping.get(code, "").strip():
                errors.append(f"Enter a full form for DOCX Qtype {code}.")

        if errors:
            message = errors[0] if len(errors) == 1 else errors[0] + f" (+{len(errors) - 1} more)"
            self.qtype_feedback_var.set(message)
            self.qtype_feedback_label.configure(fg=self.COLORS["danger"], bg="#2A1520")
            if show_errors:
                show_message("error", "Invalid Qtype configuration", "\n".join(errors), parent=self.root)
            return None

        used_text = ", ".join(self.detected_qtype_codes) if self.detected_qtype_codes else "none yet"
        self.qtype_feedback_var.set(f"Stored: {len(mapping)} • DOCX uses: {used_text}")
        self.qtype_feedback_label.configure(fg=self.COLORS["success"], bg="#0B2A25")
        return mapping

    def _schedule_qtype_refresh(self) -> None:
        if self._qtype_refreshing:
            return
        if self._qtype_refresh_job is not None:
            try:
                self.root.after_cancel(self._qtype_refresh_job)
            except Exception:
                pass
        self._qtype_refresh_job = self.root.after(250, self._refresh_qtype_state)

    def _refresh_qtype_state(self, save_if_valid: bool = True) -> None:
        self._qtype_refresh_job = None
        if self._qtype_refreshing:
            return
        self._qtype_refreshing = True
        try:
            mapping = self._collect_qtype_map(show_errors=False)
            if mapping is not None:
                # Normalize visible codes and turn completed new rows into saved rows.
                for row in self.qtype_rows:
                    code = self._qtype_row_code(row)
                    code_var = row.get("code_var")
                    if code and code_var is not None:
                        try:
                            if str(code_var.get()) != code:
                                code_var.set(code)
                        except Exception:
                            pass
                    if code in mapping and code not in DEFAULT_QTYPE_MAP:
                        row["origin"] = "saved"
                if save_if_valid:
                    try:
                        save_qtype_store(mapping)
                        self.qtype_store = OrderedDict(mapping)
                    except GenerationError as exc:
                        self.qtype_feedback_var.set(str(exc))
                        self.qtype_feedback_label.configure(fg=self.COLORS["warning"], bg="#30270E")
                else:
                    self.qtype_store = OrderedDict(mapping)
            self._refresh_qtype_row_badges()
        finally:
            self._qtype_refreshing = False
        self._refresh_generate_state()

    def _refresh_qtype_row_badges(self) -> None:
        for row in self.qtype_rows:
            code = self._qtype_row_code(row)
            origin = str(row.get("origin", "manual"))
            in_docx = bool(code and code in self.detected_qtype_codes)
            stored = bool(code and code in self.qtype_store)

            if origin == "builtin":
                status = "Built-in" + (" • DOCX" if in_docx else "")
                color = self.COLORS["blue"]
            elif in_docx and stored:
                status = "Saved • DOCX"
                color = self.COLORS["success"]
            elif in_docx:
                status = "DOCX required"
                color = self.COLORS["warning"]
            elif stored or origin == "saved":
                status = "Saved"
                color = self.COLORS["success"]
            else:
                status = "New"
                color = self.COLORS["muted"]

            label = row.get("status_label")
            if label is not None:
                try:
                    label.configure(text=status, fg=color)
                except Exception:
                    pass

            code_entry = row.get("code_entry")
            if code_entry is not None:
                desired = "normal" if origin == "manual" and not self.busy else "disabled"
                try:
                    code_entry.configure(state=desired)
                except Exception:
                    pass

            full_entry = row.get("full_form_entry")
            if full_entry is not None:
                try:
                    full_entry.configure(state="disabled" if self.busy else "normal")
                except Exception:
                    pass

            remove_button = row.get("remove_button")
            if remove_button is not None:
                desired = "disabled" if self.busy or in_docx else "normal"
                try:
                    remove_button.configure(state=desired)
                except Exception:
                    pass

    # --------------------------- Question Type management UI ---------------------------

    def _populate_question_type_rows(self) -> None:
        for row in list(self.question_type_rows):
            frame = row.get("frame")
            if frame is not None:
                try:
                    frame.destroy()
                except Exception:
                    pass
        self.question_type_rows.clear()

        for code, full_form in self.question_type_store.items():
            origin = "builtin" if code in DEFAULT_QUESTION_TYPE_MAP else "saved"
            self._add_question_type_row(code=code, full_form=full_form, origin=origin, refresh=False)
        self._refresh_question_type_state(save_if_valid=False)

    def _add_question_type_row(
        self,
        code: str = "",
        full_form: str = "",
        origin: str = "manual",
        refresh: bool = True,
    ) -> None:
        c = self.COLORS
        row_frame = self.tk.Frame(self.question_type_container, bg=c["card_alt"])
        row_frame.pack(fill="x", padx=8, pady=5)

        code_var = self.tk.StringVar(value=normalize_question_type_code(code) if code else "")
        full_form_var = self.tk.StringVar(value=str(full_form or ""))
        code_entry = self._entry(row_frame, code_var, width=10, justify="center")
        code_entry.grid(row=0, column=0, sticky="ew", padx=(0, 4), ipady=7)
        full_form_entry = self._entry(row_frame, full_form_var, width=38, justify="left")
        full_form_entry.grid(row=0, column=1, sticky="ew", padx=4, ipady=7)

        status_label = self.tk.Label(
            row_frame,
            text="New",
            width=18,
            bg=c["card_alt"],
            fg=c["muted"],
            font=self._font(8, "bold"),
        )
        status_label.grid(row=0, column=2, sticky="ew", padx=4)

        row: Dict[str, object] = {
            "frame": row_frame,
            "code_var": code_var,
            "full_form_var": full_form_var,
            "code_entry": code_entry,
            "full_form_entry": full_form_entry,
            "status_label": status_label,
            "origin": origin,
        }

        if origin == "builtin":
            self.tk.Label(
                row_frame,
                text="Built-in",
                width=8,
                bg=c["card_alt"],
                fg=c["muted"],
                font=self._font(8, "bold"),
            ).grid(row=0, column=3, sticky="e", padx=(5, 0))
        else:
            remove_button = self._button(
                row_frame,
                "Remove",
                lambda r=row: self._remove_question_type_row(r),
                kind="danger",
                width=7,
                padx=8,
                pady=7,
            )
            remove_button.grid(row=0, column=3, sticky="e", padx=(5, 0))
            row["remove_button"] = remove_button
            self.config_widgets.append(remove_button)

        row_frame.columnconfigure(1, weight=1)
        self.question_type_rows.append(row)
        self.config_widgets.extend([code_entry, full_form_entry])

        code_var.trace_add("write", lambda *_: self._schedule_question_type_refresh())
        full_form_var.trace_add("write", lambda *_: self._schedule_question_type_refresh())
        code_entry.bind("<FocusOut>", lambda _e, r=row: self._normalize_question_type_row_code(r))

        self._refresh_question_type_row_badges()
        if refresh:
            self._schedule_question_type_refresh()

    def _normalize_question_type_row_code(self, row: Dict[str, object]) -> None:
        if self._question_type_refreshing:
            return
        code_var = row.get("code_var")
        if code_var is None:
            return
        try:
            normalized = normalize_question_type_code(code_var.get())
            if normalized != str(code_var.get()):
                code_var.set(normalized)
        except Exception:
            pass

    def _question_type_row_code(self, row: Dict[str, object]) -> str:
        try:
            return normalize_question_type_code(row["code_var"].get())
        except Exception:
            return ""

    def _find_question_type_row(self, code: str) -> Optional[Dict[str, object]]:
        wanted = normalize_question_type_code(code)
        for row in self.question_type_rows:
            if self._question_type_row_code(row) == wanted:
                return row
        return None

    def _remove_question_type_row(self, row: Dict[str, object]) -> None:
        if row not in self.question_type_rows:
            return
        code = self._question_type_row_code(row)
        if code and code in self.detected_question_type_codes:
            show_message(
                "error",
                "Question Type is in use",
                f"Cannot remove Question Type {code} because it is used by the selected DOCX.",
                parent=self.root,
            )
            return
        if str(row.get("origin", "")) == "builtin":
            return

        self.question_type_rows.remove(row)
        frame = row.get("frame")
        if frame is not None:
            try:
                frame.destroy()
            except Exception:
                pass
        self._refresh_question_type_state(save_if_valid=True)

    def _sync_detected_question_types(self, codes: Sequence[str]) -> None:
        normalized: List[str] = []
        seen = set()
        for value in codes:
            code = normalize_question_type_code(value)
            if code and code not in seen:
                seen.add(code)
                normalized.append(code)
        self.detected_question_type_codes = normalized

        # Drop only temporary DOCX-only blank rows that are no longer needed.
        for row in list(self.question_type_rows):
            if str(row.get("origin", "")) != "detected":
                continue
            code = self._question_type_row_code(row)
            if code in self.detected_question_type_codes or code in self.question_type_store:
                continue
            self.question_type_rows.remove(row)
            frame = row.get("frame")
            if frame is not None:
                try:
                    frame.destroy()
                except Exception:
                    pass

        for code in self.detected_question_type_codes:
            if self._find_question_type_row(code) is None:
                self._add_question_type_row(code=code, full_form="", origin="detected", refresh=False)

        self._refresh_question_type_row_badges()
        self._refresh_question_type_state(save_if_valid=False)

    def _collect_question_type_map(self, show_errors: bool = False) -> Optional["OrderedDict[str, str]"]:
        errors: List[str] = []
        mapping: "OrderedDict[str, str]" = OrderedDict()

        for index, row in enumerate(self.question_type_rows, start=1):
            code = self._question_type_row_code(row)
            try:
                full_form = str(row["full_form_var"].get()).strip()
            except Exception:
                full_form = ""

            if not code and not full_form:
                errors.append(f"Question Type row {index}: enter a code and full form, or remove the row.")
                continue
            if not code:
                errors.append(f"Question Type row {index}: Code cannot be blank.")
                continue
            if not full_form:
                errors.append(f"Question Type {code}: Full form cannot be blank.")
                continue
            if code in mapping:
                errors.append(f"Question Type {code} already exists. Duplicate codes are not allowed.")
                continue
            mapping[code] = full_form

        for code in self.detected_question_type_codes:
            if not mapping.get(code, "").strip():
                errors.append(f"Enter a full form for DOCX Question Type {code}.")

        if errors:
            message = errors[0] if len(errors) == 1 else errors[0] + f" (+{len(errors) - 1} more)"
            self.question_type_feedback_var.set(message)
            self.question_type_feedback_label.configure(fg=self.COLORS["danger"], bg="#2A1520")
            if show_errors:
                show_message("error", "Invalid Question Type configuration", "\n".join(errors), parent=self.root)
            return None

        used_text = ", ".join(self.detected_question_type_codes) if self.detected_question_type_codes else "none yet"
        self.question_type_feedback_var.set(f"Stored: {len(mapping)} • DOCX uses: {used_text}")
        self.question_type_feedback_label.configure(fg=self.COLORS["success"], bg="#0B2A25")
        return mapping

    def _schedule_question_type_refresh(self) -> None:
        if self._question_type_refreshing:
            return
        if self._question_type_refresh_job is not None:
            try:
                self.root.after_cancel(self._question_type_refresh_job)
            except Exception:
                pass
        self._question_type_refresh_job = self.root.after(250, self._refresh_question_type_state)

    def _refresh_question_type_state(self, save_if_valid: bool = True) -> None:
        self._question_type_refresh_job = None
        if self._question_type_refreshing:
            return
        self._question_type_refreshing = True
        try:
            mapping = self._collect_question_type_map(show_errors=False)
            if mapping is not None:
                # Normalize visible codes and turn completed new rows into saved rows.
                for row in self.question_type_rows:
                    code = self._question_type_row_code(row)
                    code_var = row.get("code_var")
                    if code and code_var is not None:
                        try:
                            if str(code_var.get()) != code:
                                code_var.set(code)
                        except Exception:
                            pass
                    if code in mapping and code not in DEFAULT_QUESTION_TYPE_MAP:
                        row["origin"] = "saved"
                if save_if_valid:
                    try:
                        save_question_type_store(mapping)
                        self.question_type_store = OrderedDict(mapping)
                    except GenerationError as exc:
                        self.question_type_feedback_var.set(str(exc))
                        self.question_type_feedback_label.configure(fg=self.COLORS["warning"], bg="#30270E")
                else:
                    self.question_type_store = OrderedDict(mapping)
            self._refresh_question_type_row_badges()
        finally:
            self._question_type_refreshing = False
        self._refresh_generate_state()

    def _refresh_question_type_row_badges(self) -> None:
        for row in self.question_type_rows:
            code = self._question_type_row_code(row)
            origin = str(row.get("origin", "manual"))
            in_docx = bool(code and code in self.detected_question_type_codes)
            stored = bool(code and code in self.question_type_store)

            if origin == "builtin":
                status = "Built-in" + (" • DOCX" if in_docx else "")
                color = self.COLORS["blue"]
            elif in_docx and stored:
                status = "Saved • DOCX"
                color = self.COLORS["success"]
            elif in_docx:
                status = "DOCX required"
                color = self.COLORS["warning"]
            elif stored or origin == "saved":
                status = "Saved"
                color = self.COLORS["success"]
            else:
                status = "New"
                color = self.COLORS["muted"]

            label = row.get("status_label")
            if label is not None:
                try:
                    label.configure(text=status, fg=color)
                except Exception:
                    pass

            code_entry = row.get("code_entry")
            if code_entry is not None:
                desired = "normal" if origin == "manual" and not self.busy else "disabled"
                try:
                    code_entry.configure(state=desired)
                except Exception:
                    pass

            full_entry = row.get("full_form_entry")
            if full_entry is not None:
                try:
                    full_entry.configure(state="disabled" if self.busy else "normal")
                except Exception:
                    pass

            remove_button = row.get("remove_button")
            if remove_button is not None:
                desired = "disabled" if self.busy or in_docx else "normal"
                try:
                    remove_button.configure(state=desired)
                except Exception:
                    pass


    # --------------------------- output mode UI ---------------------------

    def _selected_output_modes(self) -> Tuple[str, ...]:
        return tuple(
            mode for mode in OUTPUT_MODE_ORDER
            if bool(self.output_mode_vars[mode].get())
        )

    def _toggle_output_mode(self, mode: str) -> None:
        if self.busy or mode not in self.output_mode_vars:
            return
        variable = self.output_mode_vars[mode]
        variable.set(not bool(variable.get()))
        self._refresh_output_mode_ui()
        selected = self._selected_output_modes()
        if self.selected_docx:
            if not selected:
                self._set_status("DOCX ready — select at least one HTM mode", "warning")
            elif (
                self._collect_attempt_settings(show_errors=False) is not None
                and self._collect_qtype_map(show_errors=False) is not None
                and self._collect_question_type_map(show_errors=False) is not None
            ):
                self._set_status("Ready — review mappings, attempts, and selected HTM modes, then generate", "success")
        self._refresh_generate_state()

    def _refresh_output_mode_ui(self) -> None:
        if not hasattr(self, "output_mode_buttons"):
            return
        c = self.COLORS
        accent_by_mode = {"self": c["success"], "test": c["blue"], "crm": c["warning"]}
        selected = self._selected_output_modes()
        for mode, button in self.output_mode_buttons.items():
            is_selected = mode in selected
            accent = accent_by_mode.get(mode, c["accent"])
            try:
                button.configure(
                    bg=accent if is_selected else c["card_alt"],
                    fg=c["bg"] if is_selected else c["text"],
                    activebackground=accent,
                    activeforeground=c["bg"],
                    relief="sunken" if is_selected else "flat",
                )
            except Exception:
                pass

        labels = [OUTPUT_MODE_LABELS[mode] for mode in selected]
        if not labels:
            self.mode_feedback_var.set("No mode selected — choose Self Learning, Test, CRM, or any combination")
            if hasattr(self, "mode_feedback_label"):
                self.mode_feedback_label.configure(fg=c["warning"], bg="#312611")
        elif len(selected) == len(OUTPUT_MODE_ORDER):
            self.mode_feedback_var.set("All 3 modes selected — full Self Learning + Test + CRM HTM")
            if hasattr(self, "mode_feedback_label"):
                self.mode_feedback_label.configure(fg=c["success"], bg="#0B2A25")
        else:
            self.mode_feedback_var.set("Generated HTM modes: " + " + ".join(labels))
            if hasattr(self, "mode_feedback_label"):
                self.mode_feedback_label.configure(fg=c["success"], bg="#0B2A25")

    # --------------------------- attempt rule UI ---------------------------

    def _reset_default_rules(self, initial: bool = False) -> None:
        question_default = 5
        test_default = 10
        initial_rules = None
        if initial:
            question_default = int(getattr(self.args, "question_attempt_max", 5) or 5)
            test_default = int(getattr(self.args, "test_attempt_max", 10) or 10)
            initial_rules = getattr(self.args, "test_attempt_rule", None)

        self.question_attempt_max_var.set(str(question_default))
        self.test_attempt_max_var.set(str(test_default))
        for row in list(self.rule_rows):
            frame = row.get("frame")
            if frame is not None:
                try:
                    frame.destroy()
                except Exception:
                    pass
        self.rule_rows.clear()

        if initial_rules:
            for rule in initial_rules:
                attempts = "Unlimited" if str(rule.attempts).casefold() == "unlimited" else str(rule.attempts)
                self._add_rule(
                    start=rule.from_value,
                    end=rule.to_value,
                    attempts=attempts,
                    refresh=False,
                )
        else:
            for start, end, attempts in self.DEFAULT_RULES:
                self._add_rule(start=start, end=end, attempts=attempts, refresh=False)

        self._reindex_rule_rows()
        self._schedule_settings_refresh()
        if not initial:
            self._set_status("Default attempt configuration restored", "success")

    def _next_rule_start(self) -> int:
        maximum = 0
        for row in self.rule_rows:
            try:
                maximum = max(maximum, int(str(row["to_var"].get()).strip()))
            except Exception:
                continue
        return maximum + 1 if maximum > 0 else 1

    def _add_rule(
        self,
        start: Optional[int] = None,
        end: Optional[int] = None,
        attempts: str = "10",
        refresh: bool = True,
    ) -> None:
        c = self.COLORS
        if start is None:
            start = self._next_rule_start()
        if end is None:
            end = start + 9

        row_frame = self.tk.Frame(self.rules_container, bg=c["card_alt"])
        row_frame.pack(fill="x", padx=8, pady=5)

        from_var = self.tk.StringVar(value=str(start))
        to_var = self.tk.StringVar(value=str(end))
        attempts_var = self.tk.StringVar(value=str(attempts))
        number_label = self.tk.Label(
            row_frame,
            text="1",
            width=5,
            bg=c["card_alt"],
            fg=c["blue"],
            font=self._font(9, "bold"),
        )
        number_label.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        from_entry = self._entry(row_frame, from_var, width=12, justify="center")
        from_entry.grid(row=0, column=1, sticky="ew", padx=4, ipady=7)
        to_entry = self._entry(row_frame, to_var, width=12, justify="center")
        to_entry.grid(row=0, column=2, sticky="ew", padx=4, ipady=7)
        attempts_entry = self._entry(row_frame, attempts_var, width=22, justify="center")
        attempts_entry.grid(row=0, column=3, sticky="ew", padx=4, ipady=7)

        row: Dict[str, object] = {
            "frame": row_frame,
            "number_label": number_label,
            "from_var": from_var,
            "to_var": to_var,
            "attempts_var": attempts_var,
            "widgets": [from_entry, to_entry, attempts_entry],
        }
        remove_button = self._button(
            row_frame,
            "×",
            lambda r=row: self._remove_rule(r),
            kind="danger",
            width=2,
            padx=8,
            pady=7,
        )
        remove_button.grid(row=0, column=4, sticky="e", padx=(5, 0))
        row["remove_button"] = remove_button
        row["widgets"].append(remove_button)  # type: ignore[union-attr]

        row_frame.columnconfigure(1, weight=1)
        row_frame.columnconfigure(2, weight=1)
        row_frame.columnconfigure(3, weight=2)
        self.rule_rows.append(row)
        self.config_widgets.extend([from_entry, to_entry, attempts_entry, remove_button])

        for variable in (from_var, to_var, attempts_var):
            variable.trace_add("write", lambda *_: self._schedule_settings_refresh())

        self._reindex_rule_rows()
        if refresh:
            self._schedule_settings_refresh()

    def _remove_rule(self, row: Dict[str, object]) -> None:
        if row not in self.rule_rows:
            return
        self.rule_rows.remove(row)
        frame = row.get("frame")
        if frame is not None:
            try:
                frame.destroy()
            except Exception:
                pass
        self._reindex_rule_rows()
        self._schedule_settings_refresh()

    def _reindex_rule_rows(self) -> None:
        for index, row in enumerate(self.rule_rows, start=1):
            label = row.get("number_label")
            if label is not None:
                try:
                    label.configure(text=str(index))
                except Exception:
                    pass

    def _collect_attempt_settings(self, show_errors: bool = False) -> Optional[AttemptSettings]:
        errors: List[str] = []
        try:
            question_max = int(self.question_attempt_max_var.get().strip())
        except Exception:
            question_max = 0
        if question_max < 1 or question_max > 99:
            errors.append("Question Attempts Max must be an integer from 1 to 99.")

        try:
            test_max = int(self.test_attempt_max_var.get().strip())
        except Exception:
            test_max = 0
        if test_max < 1 or test_max > 999:
            errors.append("Default Test Attempts Max must be an integer from 1 to 999.")

        if not self.rule_rows:
            errors.append("Add at least one Test Attempt Rule.")

        rules: List[TestAttemptRule] = []
        previous_to = 0
        for index, row in enumerate(self.rule_rows, start=1):
            try:
                from_value = int(str(row["from_var"].get()).strip())
            except Exception:
                from_value = 0
            try:
                to_value = int(str(row["to_var"].get()).strip())
            except Exception:
                to_value = 0
            attempts_raw = str(row["attempts_var"].get()).strip()

            if from_value < 1:
                errors.append(f"Rule {index}: From must be a positive integer.")
            if to_value < from_value:
                errors.append(f"Rule {index}: To must be greater than or equal to From.")
            if index > 1 and from_value <= previous_to:
                errors.append(f"Rule {index}: range overlaps or is out of order.")
            previous_to = max(previous_to, to_value)

            lowered = attempts_raw.casefold()
            if lowered in {"unlimited", "0"}:
                attempts_value: AttemptValue = "unlimited"
            else:
                try:
                    attempts_number = int(attempts_raw)
                except Exception:
                    attempts_number = 0
                if attempts_number <= 0:
                    errors.append(f"Rule {index}: Attempts must be Unlimited, 0, or a positive integer.")
                    attempts_number = 1
                attempts_value = attempts_number

            rules.append(TestAttemptRule(from_value, to_value, attempts_value))

        if errors:
            message = errors[0] if len(errors) == 1 else errors[0] + f" (+{len(errors) - 1} more)"
            self.settings_feedback_var.set(message)
            self.settings_feedback_label.configure(fg=self.COLORS["danger"], bg="#2A1520")
            if show_errors:
                show_message("error", "Invalid attempt configuration", "\n".join(errors), parent=self.root)
            return None

        settings = AttemptSettings(
            question_attempt_max=question_max,
            test_attempt_max=test_max,
            test_attempt_rules=tuple(rules),
        )
        self.settings_feedback_var.set(settings.summary())
        self.settings_feedback_label.configure(fg=self.COLORS["success"], bg="#0B2A25")
        return settings

    def _schedule_settings_refresh(self) -> None:
        if self._settings_refresh_job is not None:
            try:
                self.root.after_cancel(self._settings_refresh_job)
            except Exception:
                pass
        self._settings_refresh_job = self.root.after(120, self._refresh_settings_state)

    def _refresh_settings_state(self) -> None:
        self._settings_refresh_job = None
        self._collect_attempt_settings(show_errors=False)
        self._refresh_generate_state()

    # --------------------------- file selection ---------------------------

    def _initial_select(self) -> None:
        if getattr(self.args, "docx", None):
            path = Path(self.args.docx).expanduser().resolve()
            self._accept_docx(path, auto_start=False)
        else:
            self.select_docx(auto_start=False)

    def select_docx(self, auto_start: bool = False) -> None:
        if self.worker and self.worker.is_alive():
            return
        selected = choose_docx_gui(parent=self.root)
        if selected is None:
            self._set_status("No DOCX selected", "warning")
            return
        self._accept_docx(selected.expanduser().resolve(), auto_start=auto_start)

    def _accept_docx(self, docx_path: Path, auto_start: bool) -> None:
        if not docx_path.is_file() or docx_path.suffix.lower() != ".docx":
            show_message("error", APP_NAME, f"Select a valid .docx file:\n{docx_path}", parent=self.root)
            return

        self.selected_docx = docx_path
        self.docx_var.set(str(docx_path))
        try:
            preview_tables = read_docx_tables(docx_path)
            preview_table = select_question_table(preview_tables)
            preview_questions = parse_questions(preview_table)
            detected_qtypes = qtype_codes_from_questions(preview_questions)
            detected_question_types = question_type_codes_from_questions(preview_questions)
            self.selected_question_count = len(preview_questions)
            self.question_count_var.set(f"Total scored questions: {self.selected_question_count}")
        except GenerationError as exc:
            self.selected_docx = None
            self.selected_question_count = 0
            self.question_count_var.set("Total scored questions: —")
            self._set_status("DOCX could not be parsed", "danger")
            show_message("error", APP_NAME, str(exc), parent=self.root)
            self._refresh_generate_state()
            return
        self._sync_detected_qtypes(detected_qtypes)
        self._sync_detected_question_types(detected_question_types)
        base = docx_path.parent

        configured_q = getattr(self.args, "images_q", None)
        configured_e = getattr(self.args, "images_e", None)
        images_q = Path(configured_q).expanduser().resolve() if configured_q else locate_named_directory(base, "imagesQ")
        images_e = Path(configured_e).expanduser().resolve() if configured_e else locate_named_directory(base, "imagesE")

        if images_q is None or not images_q.is_dir():
            images_q = choose_directory_gui("Select the imagesQ folder", base, parent=self.root)
        if images_e is None or not images_e.is_dir():
            images_e = choose_directory_gui("Select the imagesE folder", base, parent=self.root)

        if images_q is None or not images_q.is_dir():
            self.resources_var.set("imagesQ folder not selected")
            self._set_status("Select a valid imagesQ folder", "danger")
            self._refresh_generate_state()
            return
        if images_e is None or not images_e.is_dir():
            self.resources_var.set("imagesE folder not selected")
            self._set_status("Select a valid imagesE folder", "danger")
            self._refresh_generate_state()
            return

        self.images_q = images_q.resolve()
        self.images_e = images_e.resolve()
        configured_output = getattr(self.args, "output", None)
        self.output_path = (
            Path(configured_output).expanduser().resolve()
            if configured_output
            else default_output_path(docx_path)
        )

        self.resources_var.set(f"imagesQ: {self.images_q}\nimagesE: {self.images_e}")
        self.output_var.set(str(self.output_path))
        if self._collect_qtype_map(show_errors=False) is None:
            self._set_status("Complete the required Qtype full forms before generating", "warning")
        elif self._collect_question_type_map(show_errors=False) is None:
            self._set_status("Complete the required Question Type full forms before generating", "warning")
        elif not self._selected_output_modes():
            self._set_status("DOCX ready — select at least one HTM mode", "warning")
        else:
            self._set_status("Ready — review mappings, attempts, and selected HTM modes, then generate", "success")
        self._refresh_output_mode_ui()
        self._refresh_generate_state()

        if auto_start:
            self.root.after(100, self.start_generation)

    # --------------------------- generation ---------------------------

    def _set_status(self, text: str, tone: str = "info") -> None:
        self.status_var.set(text)
        color = {
            "info": self.COLORS["blue"],
            "success": self.COLORS["success"],
            "warning": self.COLORS["warning"],
            "danger": self.COLORS["danger"],
        }.get(tone, self.COLORS["blue"])
        if hasattr(self, "status_dot"):
            self.status_dot.configure(fg=color)

    def _refresh_generate_state(self) -> None:
        if not hasattr(self, "generate_button"):
            return
        settings_ok = self._collect_attempt_settings(show_errors=False) is not None
        qtypes_ok = self._collect_qtype_map(show_errors=False) is not None
        question_types_ok = self._collect_question_type_map(show_errors=False) is not None
        modes_ok = bool(self._selected_output_modes())
        ready = bool(
            self.selected_docx
            and self.images_q
            and self.images_e
            and self.output_path
            and settings_ok
            and qtypes_ok
            and question_types_ok
            and modes_ok
            and not self.busy
        )
        self.generate_button.configure(state="normal" if ready else "disabled")

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.select_button.configure(state=state)
        for widget in self.config_widgets:
            try:
                widget.configure(state=state)
            except Exception:
                pass
        if busy:
            self.generate_button.configure(state="disabled")
            self.progress.start(12)
        else:
            self.progress.stop()
        self._refresh_qtype_row_badges()
        self._refresh_question_type_row_badges()
        self._refresh_output_mode_ui()
        if not busy:
            self._refresh_generate_state()

    def start_generation(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        if not self.selected_docx or not self.images_q or not self.images_e or not self.output_path:
            self.select_docx(auto_start=False)
            return

        attempt_settings = self._collect_attempt_settings(show_errors=True)
        if attempt_settings is None:
            return
        qtype_map = self._collect_qtype_map(show_errors=True)
        if qtype_map is None:
            return
        question_type_map = self._collect_question_type_map(show_errors=True)
        if question_type_map is None:
            return
        enabled_modes = self._selected_output_modes()
        if not enabled_modes:
            show_message("error", "Select HTM mode", "Select at least one mode: Self Learning, Test, or CRM.", parent=self.root)
            self._refresh_generate_state()
            return
        try:
            save_qtype_store(qtype_map)
            self.qtype_store = OrderedDict(qtype_map)
            save_question_type_store(question_type_map)
            self.question_type_store = OrderedDict(question_type_map)
        except GenerationError as exc:
            show_message("error", "Vocabulary storage error", str(exc), parent=self.root)
            return

        docx_path = self.selected_docx
        images_q = self.images_q
        images_e = self.images_e
        output_path = self.output_path
        mode_text = " + ".join(OUTPUT_MODE_LABELS[mode] for mode in enabled_modes)
        self._set_status(f"Processing DOCX for {mode_text}…", "info")
        self._set_busy(True)

        def worker() -> None:
            try:
                report = generate(
                    docx_path=docx_path,
                    output_path=output_path,
                    images_q=images_q,
                    images_e=images_e,
                    attempt_settings=attempt_settings,
                    qtype_map=qtype_map,
                    question_type_map=question_type_map,
                    enabled_modes=enabled_modes,
                )
                self.result_queue.put(("success", (report, output_path, attempt_settings, enabled_modes)))
            except GenerationError as exc:
                self.result_queue.put(("generation_error", str(exc)))
            except Exception as exc:
                self.result_queue.put(("unexpected_error", (str(exc), traceback.format_exc())))

        self.worker = threading.Thread(target=worker, name="skillnox-generator", daemon=True)
        self.worker.start()

    def _poll_results(self) -> None:
        try:
            while True:
                kind, payload = self.result_queue.get_nowait()
                if kind == "success":
                    report, output_path, settings, enabled_modes = payload  # type: ignore[misc]
                    self._handle_success(report, output_path, settings, enabled_modes)
                elif kind == "generation_error":
                    self._handle_error(str(payload))
                elif kind == "unexpected_error":
                    error, details = payload  # type: ignore[misc]
                    self._handle_error(f"Unexpected error: {error}\n\n{details}")
        except queue.Empty:
            pass
        finally:
            if self.root.winfo_exists():
                self.root.after(100, self._poll_results)

    def _handle_success(
        self,
        report: GenerationReport,
        output_path: Path,
        settings: AttemptSettings,
        enabled_modes: Sequence[str],
    ) -> None:
        self._set_busy(False)
        mode_text = "Reference shell mode" if report.exact_reference_mode else "Generic DOCX mode"
        should_open = not bool(getattr(self.args, "no_open", False))
        if should_open:
            opened_in_chrome, browser_text = open_html_in_chrome(output_path)
        else:
            opened_in_chrome = False
            browser_text = "Automatic browser opening was disabled."

        rules_text = ", ".join(
            f"{rule.from_value}–{rule.to_value}: "
            + ("Unlimited" if str(rule.attempts).casefold() == "unlimited" else str(rule.attempts))
            for rule in settings.test_attempt_rules
        )
        selected_mode_text = " + ".join(OUTPUT_MODE_LABELS.get(mode, mode) for mode in enabled_modes)
        message = (
            f"Generated successfully.\n\n"
            f"Output: {output_path}\n"
            f"HTM modes: {selected_mode_text}\n"
            f"Questions: {report.questions}\n"
            f"Options: {report.options}\n"
            f"QSVG references: {report.qsvg_references}\n"
            f"ESVG references: {report.esvg_references}\n"
            f"Qtypes (P/T): {', '.join(report.qtype_codes)}\n"
            f"Question Types (QTYPE): {', '.join(report.question_type_codes)}\n"
            f"Question Attempts Max: {settings.question_attempt_max}\n"
            f"Default Test Attempts Max: {settings.test_attempt_max}\n"
            f"Test Attempt Rules: {rules_text}\n"
            f"Mode: {mode_text}\n"
            f"SVG comparison: SKIPPED\n"
            f"SHA-256: {report.output_sha256}\n\n"
            f"{browser_text}"
        )
        if report.warnings:
            message += "\n\nWarnings:\n- " + "\n- ".join(report.warnings[:12])

        self._set_status("Completed — opened in Chrome" if opened_in_chrome else "Completed", "success")
        show_message("info", APP_NAME, message, parent=self.root)

    def _handle_error(self, message: str) -> None:
        self._set_busy(False)
        self.exit_code = 2
        self._set_status("Generation failed", "danger")
        show_message("error", APP_NAME, message, parent=self.root)

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            _, _, _, messagebox = _tk_modules()
            if messagebox is not None:
                close_now = messagebox.askyesno(
                    APP_NAME,
                    "Generation is still running. Close the program?",
                    parent=self.root,
                )
                if not close_now:
                    return
        try:
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        except Exception:
            pass
        self.root.destroy()

def launch_gui(args) -> int:
    _set_windows_dpi_awareness()
    tk, _, _, _ = _tk_modules()
    if tk is None:
        print("ERROR: Tkinter is not available. Run the script with a DOCX path from the command line.", file=sys.stderr)
        return 2
    root = tk.Tk()
    app = SkillNoxGeneratorGUI(root, args)
    root.mainloop()
    return app.exit_code


# ---------------------------------------------------------------------------
# DOCX/OOXML reader
# ---------------------------------------------------------------------------

def _paragraph_text(paragraph: ET.Element) -> str:
    out: List[str] = []
    for node in paragraph.iter():
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag in (W + "br", W + "cr"):
            out.append("\n")
        elif node.tag == W + "tab":
            out.append("\t")
    return "".join(out)


def _cell_text(cell: ET.Element) -> str:
    paragraphs = cell.findall(".//" + W + "p")
    return "\n".join(_paragraph_text(p) for p in paragraphs).replace("\r\n", "\n").replace("\r", "\n").strip()


def read_docx_tables(docx_path: Path) -> List[List[List[str]]]:
    try:
        with zipfile.ZipFile(docx_path) as archive:
            document_xml = archive.read("word/document.xml")
    except (OSError, KeyError, zipfile.BadZipFile) as exc:
        raise GenerationError(f"Cannot read DOCX file: {exc}") from exc

    try:
        root = ET.fromstring(document_xml)
    except ET.ParseError as exc:
        raise GenerationError(f"Invalid Word XML: {exc}") from exc

    tables: List[List[List[str]]] = []
    for tbl in root.findall(".//" + W + "tbl"):
        rows: List[List[str]] = []
        for tr in tbl.findall("./" + W + "tr"):
            cells = [_cell_text(tc) for tc in tr.findall("./" + W + "tc")]
            rows.append(cells)
        if rows:
            tables.append(rows)
    return tables


def normalize_header(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", value.upper())


def select_question_table(tables: Sequence[List[List[str]]]) -> List[List[str]]:
    for table in tables:
        if not table:
            continue
        headers = {normalize_header(v) for v in table[0]}
        if "CCODE" in headers and "SOLUTION" in headers and "SUBTOPIC" in headers:
            return table
    raise GenerationError("No SkillNox question table was found. Required headers include C-CODE, SOLUTION and SUBTOPIC.")


def fingerprint_table(table: Sequence[Sequence[str]]) -> str:
    normalized = [[str(cell).replace("\r\n", "\n").replace("\r", "\n").strip() for cell in row] for row in table]
    payload = json.dumps(normalized, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _header_index(header: Sequence[str], name: str, fallback: int) -> int:
    wanted = normalize_header(name)
    for i, value in enumerate(header):
        if normalize_header(value) == wanted:
            return i
    return fallback


def strip_subtopic_code(text: str) -> str:
    value = text.strip()
    value = re.sub(r"^\s*\d+(?:\.\d+)+\s*(?:[-–—:]\s*)?", "", value)
    return value.strip() or text.strip()


def normalize_feature_qid(raw: str) -> str:
    digits = "".join(re.findall(r"\d+", raw))
    return digits or re.sub(r"\W+", "", raw)


SPECIAL_QUESTION_TYPES = {"TRFL", "BLNK", "DRWD", "MTCH"}
SUPPORTED_QUESTION_TYPES = {"MCQ", "MCQS"} | SPECIAL_QUESTION_TYPES
SPECIAL_SOLUTION_MARK_RE = re.compile(r"^\s*Solution\s*:\s*", re.IGNORECASE)
INLINE_ANSWER_MARK_RE = re.compile(r"\*([^*\r\n]+)\*")


def _normalise_special_answer(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


CASE_STUDY_PARENT_RE = re.compile(
    r"^\s*CASE\s+STUDY\s*/\s*PASSAGE\s+BASED\s+QUESTION\s*:\s*(.*)$",
    re.IGNORECASE,
)
CASE_STUDY_ROMAN_RE = re.compile(
    r"^\s*\((i|ii|iii|iv|v|vi|vii|viii|ix|x)\)\s*",
    re.IGNORECASE,
)


def _parse_case_study_parent(raw: str) -> Optional[Tuple[str, List[str]]]:
    lines = _trim_special_lines(raw.replace("\r\n", "\n").replace("\r", "\n").split("\n"))
    if not lines:
        return None
    source_number = ""
    first = lines[0]
    source_match = SOURCE_NUMBER_RE.match(first)
    if source_match:
        source_number = source_match.group(1).strip()
        first = SOURCE_NUMBER_RE.sub("", first, count=1)
    match = CASE_STUDY_PARENT_RE.match(first)
    if not match:
        return None
    passage_lines: List[str] = []
    inline = match.group(1).strip()
    if inline:
        passage_lines.append(inline)
    passage_lines.extend(lines[1:])
    passage_lines = _trim_special_lines(passage_lines)
    if not passage_lines:
        raise GenerationError("Case Study parent passage is empty.")
    return source_number, passage_lines


def _strip_case_study_child_prefix(raw: str) -> Tuple[str, str]:
    lines = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    first_index = next((i for i, line in enumerate(lines) if line.strip()), -1)
    if first_index < 0:
        return "", raw
    line = lines[first_index]
    source_prefix = ""
    source_match = SOURCE_NUMBER_RE.match(line)
    if source_match:
        source_prefix = f"[{source_match.group(1).strip()}] "
        line = SOURCE_NUMBER_RE.sub("", line, count=1)
    roman = CASE_STUDY_ROMAN_RE.match(line)
    if not roman:
        return "", raw
    part = roman.group(1).lower()
    lines[first_index] = source_prefix + line[roman.end():]
    return part, "\n".join(lines)


def _trim_special_lines(lines: Sequence[str]) -> List[str]:
    values = [str(line).strip() for line in lines]
    while values and not values[0]:
        values.pop(0)
    while values and not values[-1]:
        values.pop()
    return values


def _remove_special_instruction_line(lines: List[str], qtype: str) -> List[str]:
    if not lines:
        return lines
    instruction_aliases = {
        "TRFL": {"answer true or false", "true or false", "true/false", "true false"},
        "BLNK": {"fill in the blank", "fill in the blanks"},
        "DRWD": {"drag the word", "drag the words"},
        "MTCH": {"match the following", "match following"},
    }
    first = re.sub(r"\s+", " ", lines[0]).strip().casefold()
    if first in instruction_aliases.get(qtype, set()):
        return _trim_special_lines(lines[1:])
    return lines


def _accepted_answer_values(spec: str) -> List[str]:
    values: List[str] = []
    seen = set()
    for raw in str(spec or "").split("/"):
        value = re.sub(r"\s+", " ", raw.strip())
        key = value.casefold()
        if value and key not in seen:
            values.append(value)
            seen.add(key)
    return values


def parse_special_ccode_cell(
    raw: str,
    row_number: int,
    qtype: str,
) -> Tuple[str, List[str], List[OptionData], str, List[str], List[str], List[Tuple[str, str]]]:
    """Parse TRFL/BLNK/DRWD/MTCH without changing the existing MCQ parser."""
    lines = [line.rstrip() for line in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    lines = _trim_special_lines(lines)
    if not lines:
        raise GenerationError(f"Row {row_number}: C-CODE cell is empty.")

    source_number = ""
    first_match = SOURCE_NUMBER_RE.match(lines[0])
    if first_match:
        source_number = first_match.group(1).strip()
        lines[0] = SOURCE_NUMBER_RE.sub("", lines[0], count=1)

    solution_index = -1
    for i, line in enumerate(lines):
        if SPECIAL_SOLUTION_MARK_RE.match(line):
            solution_index = i
            break
    if solution_index < 0:
        raise GenerationError(f"Row {row_number}: no 'Solution:' marker was found in C-CODE.")

    prompt_lines = _remove_special_instruction_line(_trim_special_lines(lines[:solution_index]), qtype)
    solution_lines = [line.strip() for line in lines[solution_index + 1:] if line.strip()]
    options: List[OptionData] = []
    answer_specs: List[str] = []
    match_pairs: List[Tuple[str, str]] = []

    if qtype == "TRFL":
        if not prompt_lines:
            raise GenerationError(f"Row {row_number}: TRFL question text is missing.")
        answer = prompt_lines[-1].strip()
        normalized = answer.casefold()
        if normalized not in {"true", "false"}:
            raise GenerationError(
                f"Row {row_number}: TRFL answer must be the final True/False line before Solution:."
            )
        question_lines = _trim_special_lines(prompt_lines[:-1])
        if not question_lines:
            raise GenerationError(f"Row {row_number}: TRFL question text is missing before the answer line.")
        answer_specs = ["True" if normalized == "true" else "False"]
        correct = normalized.upper()
        return source_number, question_lines, options, correct, solution_lines, answer_specs, match_pairs

    marker_count = sum(len(INLINE_ANSWER_MARK_RE.findall(line)) for line in prompt_lines)
    if marker_count < 1:
        raise GenerationError(f"Row {row_number}: {qtype} requires at least one *answer* marker before Solution:.")

    if qtype in {"BLNK", "DRWD"}:
        for line in prompt_lines:
            answer_specs.extend(match.group(1).strip() for match in INLINE_ANSWER_MARK_RE.finditer(line))
        if qtype == "BLNK":
            canonical: List[str] = []
            for spec in answer_specs:
                accepted = _accepted_answer_values(spec)
                if not accepted:
                    raise GenerationError(f"Row {row_number}: BLNK contains an empty *answer* marker.")
                canonical.append(_normalise_special_answer(accepted[0]))
            correct = "|".join(canonical)
        else:
            correct = "|".join(_normalise_special_answer(spec) for spec in answer_specs)
        return source_number, prompt_lines, options, correct, solution_lines, answer_specs, match_pairs

    for line in prompt_lines:
        matches = list(INLINE_ANSWER_MARK_RE.finditer(line))
        if not matches:
            continue
        if len(matches) != 1:
            raise GenerationError(
                f"Row {row_number}: each MTCH line must contain exactly one *answer* marker."
            )
        match = matches[0]
        left = (line[:match.start()] + line[match.end():]).strip()
        right = match.group(1).strip()
        if not left or not right:
            raise GenerationError(f"Row {row_number}: MTCH contains an empty left or right value.")
        match_pairs.append((left, right))
        answer_specs.append(right)
    if not match_pairs:
        raise GenerationError(f"Row {row_number}: MTCH contains no usable matching pairs.")
    correct = "|".join(_normalise_special_answer(value) for _, value in match_pairs)
    return source_number, [], options, correct, solution_lines, answer_specs, match_pairs


def parse_ccode_cell(raw: str, row_number: int) -> Tuple[str, List[str], List[OptionData], str, List[str]]:
    lines = [line.rstrip() for line in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    if not lines:
        raise GenerationError(f"Row {row_number}: C-CODE cell is empty.")

    source_number = ""
    first_match = SOURCE_NUMBER_RE.match(lines[0])
    if first_match:
        source_number = first_match.group(1).strip()
        lines[0] = SOURCE_NUMBER_RE.sub("", lines[0], count=1)

    solution_index = -1
    solution_letter = ""
    for i, line in enumerate(lines):
        match = SOLUTION_MARK_RE.match(line)
        if match:
            solution_index = i
            solution_letter = match.group(1).upper()
            break
    if solution_index < 0:
        raise GenerationError(f"Row {row_number}: no 'Solution: (x)' marker was found in C-CODE.")

    prompt_and_options = lines[:solution_index]
    solution_lines = [line for line in lines[solution_index + 1:] if line.strip()]

    question_lines: List[str] = []
    options: List[OptionData] = []
    active: Optional[OptionData] = None
    for line in prompt_and_options:
        match = OPTION_START_RE.match(line)
        if match:
            active = OptionData(letter=match.group(2).upper(), lines=[match.group(3)] if match.group(3) else [], starred=bool(match.group(1)))
            options.append(active)
        elif active is None:
            question_lines.append(line)
        else:
            active.lines.append(line)

    if not options:
        raise GenerationError(f"Row {row_number}: no answer options were found.")

    starred = [opt.letter for opt in options if opt.starred]
    correct = solution_letter or (starred[0] if starred else "")
    if starred and solution_letter and solution_letter not in starred:
        raise GenerationError(
            f"Row {row_number}: starred option {starred} does not agree with Solution: ({solution_letter.lower()})."
        )
    if not correct:
        raise GenerationError(f"Row {row_number}: the correct option could not be identified.")

    return source_number, question_lines, options, correct, solution_lines


def parse_questions(table: List[List[str]]) -> List[QuestionData]:
    header = table[0]
    indexes = {
        "slno": _header_index(header, "SLNO", 0),
        "ccode": _header_index(header, "C-CODE", 1),
        "solution": _header_index(header, "SOLUTION", 2),
        "qtype": _header_index(header, "QTYPE", 3),
        "subtopic": _header_index(header, "SUBTOPIC", 4),
        "level": _header_index(header, "LEVEL", 5),
        "pt": _header_index(header, "P/T", 6),
        "marks": _header_index(header, "MRKS", 7),
        "occur": _header_index(header, "OCCUR", 8),
        "qid": _header_index(header, "QID", 9),
        "se": _header_index(header, "S/E", 10),
    }

    questions: List[QuestionData] = []
    pending_case: Optional[Dict[str, object]] = None

    def row_cell(row: List[str], key: str) -> str:
        idx = indexes[key]
        return row[idx].strip() if idx < len(row) else ""

    def finalize_case() -> None:
        nonlocal pending_case
        if not pending_case:
            return
        children = list(pending_case.get("children", []))
        if not children:
            raise GenerationError(
                f"Case Study parent QID {pending_case.get('parent_qid_raw') or '(blank)'} has no Roman child questions."
            )
        parts = [q.case_part for q in children]
        for child in children:
            child.case_parts = list(parts)
        pending_case = None

    def append_scoreable(row: List[str], table_row_no: int, ccode: str, part: str = "") -> QuestionData:
        question_type = normalize_question_type_code(row_cell(row, "qtype")) or "MCQS"
        if question_type not in SUPPORTED_QUESTION_TYPES:
            raise GenerationError(
                f"Row {table_row_no}: unsupported QTYPE {question_type!r}. "
                "Supported codes are MCQ, MCQS, TRFL, BLNK, DRWD and MTCH."
            )
        answer_specs: List[str] = []
        match_pairs: List[Tuple[str, str]] = []
        if question_type in SPECIAL_QUESTION_TYPES:
            (
                source_number,
                question_lines,
                options,
                correct,
                solution_lines,
                answer_specs,
                match_pairs,
            ) = parse_special_ccode_cell(ccode, table_row_no, question_type)
        else:
            # MCQ/MCQS still uses the original MCQ parser unchanged.
            source_number, question_lines, options, correct, solution_lines = parse_ccode_cell(ccode, table_row_no)

        index = len(questions) + 1
        subtopic_raw = row_cell(row, "subtopic")
        qid_raw = row_cell(row, "qid")
        question = QuestionData(
            index=index,
            source_number=source_number,
            question_lines=question_lines,
            options=options,
            correct=correct,
            solution_lines=solution_lines,
            explanation_raw=row_cell(row, "solution"),
            qtype=question_type,
            subtopic_raw=subtopic_raw,
            subtopic=strip_subtopic_code(subtopic_raw).upper(),
            level=row_cell(row, "level").upper() or "MEDIUM",
            pt=row_cell(row, "pt").upper() or "A",
            marks=row_cell(row, "marks"),
            occurrence=row_cell(row, "occur"),
            qid_raw=qid_raw,
            qid=normalize_feature_qid(qid_raw),
            se=row_cell(row, "se"),
            source_row=list(row),
            answer_specs=answer_specs,
            match_pairs=match_pairs,
        )
        if pending_case and part:
            question.case_study_id = str(pending_case["id"])
            question.case_parent_qid = str(pending_case["parent_qid"])
            question.case_parent_subtopic = str(pending_case["parent_subtopic"])
            question.case_passage_lines = list(pending_case["passage_lines"])
            question.case_part = part
            question.case_question_number = int(pending_case["question_number"])
            pending_case["children"].append(question)
        questions.append(question)
        return question

    for table_row_no, row in enumerate(table[1:], start=2):
        if not any(cell.strip() for cell in row):
            continue
        ccode = row_cell(row, "ccode")
        if not ccode:
            continue

        parent = _parse_case_study_parent(ccode)
        if parent is not None:
            finalize_case()
            _, passage_lines = parent
            question_number = len(questions) + 1
            parent_qid_raw = row_cell(row, "qid")
            parent_subtopic_raw = row_cell(row, "subtopic")
            pending_case = {
                "id": f"case{question_number}",
                "question_number": question_number,
                "parent_qid_raw": parent_qid_raw,
                "parent_qid": normalize_feature_qid(parent_qid_raw),
                "parent_subtopic": strip_subtopic_code(parent_subtopic_raw).upper(),
                "passage_lines": passage_lines,
                "children": [],
            }
            continue

        if pending_case:
            part, child_ccode = _strip_case_study_child_prefix(ccode)
            if part:
                append_scoreable(row, table_row_no, child_ccode, part=part)
                continue
            finalize_case()

        append_scoreable(row, table_row_no, ccode)

    finalize_case()
    if not questions:
        raise GenerationError("The question table contains no usable question rows.")
    return questions


def locate_named_directory(base: Path, expected: str) -> Optional[Path]:
    direct = base / expected
    if direct.is_dir():
        return direct
    expected_fold = expected.casefold()
    try:
        for child in base.iterdir():
            if child.is_dir() and child.name.casefold() == expected_fold:
                return child
    except OSError:
        pass
    return None


class SvgFolder:
    def __init__(self, path: Path, label: str):
        self.path = path
        self.label = label
        self._by_name: Dict[str, Path] = {}
        for svg_path in path.rglob("*.svg"):
            key = svg_path.name.casefold()
            if key not in self._by_name:
                self._by_name[key] = svg_path

    def resolve(self, filename: str) -> Path:
        candidate = self.path / filename
        if candidate.is_file():
            return candidate
        found = self._by_name.get(Path(filename).name.casefold())
        if found:
            return found
        raise GenerationError(f"Missing {self.label} SVG: {filename}\nExpected under: {self.path}")

    def read(self, filename: str) -> str:
        path = self.resolve(filename)
        try:
            return path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            return path.read_text(encoding="utf-8", errors="replace")
        except OSError as exc:
            raise GenerationError(f"Cannot read SVG '{path}': {exc}") from exc


def extract_svg_element(svg_text: str) -> str:
    text = svg_text.lstrip("\ufeff").strip()
    start = re.search(r"<svg\b", text, re.IGNORECASE)
    if not start:
        raise GenerationError("The SVG file does not contain an <svg> root element.")
    end_matches = list(re.finditer(r"</svg\s*>", text, re.IGNORECASE))
    if not end_matches:
        raise GenerationError("The SVG file does not contain a closing </svg> element.")
    return text[start.start():end_matches[-1].end()]


def _is_hex_colour_token(value: str) -> bool:
    return len(value) in (3, 4, 6, 8) and bool(re.fullmatch(r"[0-9A-Fa-f]+", value))


def scope_svg_ids(svg_text: str, prefix: str) -> str:
    svg = extract_svg_element(svg_text)
    id_values = re.findall(r"\bid\s*=\s*([\"'])(.*?)\1", svg, flags=re.IGNORECASE | re.DOTALL)
    ids = [value for _, value in id_values if value]
    if not ids:
        return svg

    id_map: Dict[str, str] = {}
    for old in ids:
        id_map[old] = old if old.startswith(prefix) else prefix + old

    def id_attr_repl(match: re.Match[str]) -> str:
        quote = match.group(1)
        old = match.group(2)
        return f'id={quote}{id_map.get(old, old)}{quote}'

    svg = re.sub(r"\bid\s*=\s*([\"'])(.*?)\1", id_attr_repl, svg, flags=re.IGNORECASE | re.DOTALL)

    for old in sorted(id_map, key=len, reverse=True):
        new = id_map[old]
        if old == new:
            continue
        escaped = re.escape(old)
        svg = re.sub(rf"url\(\s*#\s*{escaped}\s*\)", f"url(#{new})", svg)
        svg = re.sub(rf"((?:xlink:)?href\s*=\s*[\"'])#{escaped}([\"'])", rf"\1#{new}\2", svg, flags=re.IGNORECASE)
        svg = re.sub(rf"\b(begin|end)\s*=\s*([\"']){escaped}(\.[^\"']+)([\"'])", rf"\1=\2{new}\3\4", svg, flags=re.IGNORECASE)
        if not _is_hex_colour_token(old):
            svg = re.sub(rf"#{escaped}(?=[\s\.,:>+~\[\]{{}}]|$)", f"#{new}", svg)

    def aria_repl(match: re.Match[str]) -> str:
        quote = match.group(1)
        tokens = match.group(2).split()
        return f'aria-labelledby={quote}{" ".join(id_map.get(token, token) for token in tokens)}{quote}'

    svg = re.sub(r"aria-labelledby\s*=\s*([\"'])(.*?)\1", aria_repl, svg, flags=re.IGNORECASE | re.DOTALL)
    return svg


def canonical_svg(svg_text: str) -> str:
    svg = extract_svg_element(svg_text)
    try:
        root = ET.fromstring(svg)
        for node in root.iter():
            if node.text is not None and not node.text.strip():
                node.text = None
            if node.tail is not None and not node.tail.strip():
                node.tail = None
        raw = ET.tostring(root, encoding="unicode")
        try:
            return ET.canonicalize(raw, strip_text=False)  # type: ignore[attr-defined]
        except Exception:
            return re.sub(r">\s+<", "><", raw).strip()
    except ET.ParseError:
        return re.sub(r"\s+", " ", svg).strip()


class SvgRegistry:
    def __init__(self, q_folder: SvgFolder, e_folder: SvgFolder):
        self.q_folder = q_folder
        self.e_folder = e_folder
        self.counter = 0
        self.used_q: List[str] = []
        self.used_e: List[str] = []

    def next_svg(self, filename: str, kind: str) -> Tuple[int, str]:
        self.counter += 1
        folder = self.q_folder if kind == "Q" else self.e_folder
        raw = folder.read(filename)
        scoped = scope_svg_ids(raw, f"svg{self.counter}_")
        if kind == "Q":
            self.used_q.append(filename)
        else:
            self.used_e.append(filename)
        return self.counter, scoped

    def wrapper(self, filename: str, kind: str) -> str:
        number, svg = self.next_svg(filename, kind)
        wrapper_class = "qsvg-wrapper" if kind == "Q" else "esvg-wrapper"
        align = "left" if kind == "Q" else "center"
        return (
            f'<div class="q-figure hash-img-wrap {wrapper_class}" style="margin: 16px 0; text-align: {align};">\n'
            f'<div id="svg-container-svg{number}" style="overflow: visible; padding-bottom: 12px;">\n'
            f'{svg}\n'
            f'</div>\n'
            f'</div>'
        )


# ---------------------------------------------------------------------------
# Content renderer used for non-reference DOCX files
# ---------------------------------------------------------------------------

def escape_plain(value: str) -> str:
    return html_lib.escape(value, quote=False)


def render_inline_math(content: str) -> str:
    value = content.strip()
    value = re.sub(r"^\\displaystyle\s*", "", value)
    if "\\frac" in value or "\\dfrac" in value or "\\tfrac" in value:
        return f'<span class="inline-math-frac">\\(\\displaystyle {value}\\)</span>'
    return f"\\({value}\\)"


def render_inline_text(text: str) -> str:
    out: List[str] = []
    cursor = 0
    for match in re.finditer(r"\\\((.*?)\\\)", text, flags=re.DOTALL):
        out.append(escape_plain(text[cursor:match.start()]))
        out.append(render_inline_math(match.group(1)))
        cursor = match.end()
    out.append(escape_plain(text[cursor:]))
    return "".join(out)


def render_line_with_svgs(line: str, kind: str, registry: SvgRegistry) -> str:
    out: List[str] = []
    cursor = 0
    for match in SVG_TOKEN_RE.finditer(line):
        out.append(render_inline_text(line[cursor:match.start()]))
        out.append(registry.wrapper(match.group(1).strip(), kind))
        cursor = match.end()
    out.append(render_inline_text(line[cursor:]))
    return "".join(out)


def render_lines(lines: Sequence[str], kind: str, registry: SvgRegistry, joiner: str = "<br />") -> str:
    return joiner.join(render_line_with_svgs(line, kind, registry) for line in lines if line is not None)


def _find_matching_brace(value: str, open_index: int) -> int:
    depth = 0
    for i in range(open_index, len(value)):
        if value[i] == "{":
            depth += 1
        elif value[i] == "}":
            depth -= 1
            if depth == 0:
                return i
    return -1


def _find_top_level_text_command(value: str, start_index: int = 0) -> int:
    """Return the next top-level ``\text{`` command, or -1.

    A ``\text{...}`` nested inside another TeX brace group (for example the
    numerator/denominator of ``\frac``) is part of the mathematical expression
    and must remain inside the same MathJax delimiter.  Older generator versions
    extracted every ``\text`` command to HTML, which split valid expressions such
    as ``\frac{\text{Opposite}}{\text{Hypotenuse}}`` into multiple invalid
    MathJax fragments.
    """
    depth = 0
    i = max(0, start_index)
    token = r"\text{"
    while i < len(value):
        ch = value[i]
        if ch == "\\":
            if value.startswith(token, i) and depth == 0:
                return i
            # Skip the command character plus the next character so escaped
            # braces such as \{ and \} do not affect grouping depth.
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}" and depth > 0:
            depth -= 1
        i += 1
    return -1


def latex_mixed_to_html(value: str) -> str:
    value = value.strip()
    if not value:
        return ""
    pieces: List[Tuple[str, str]] = []
    cursor = 0
    while cursor < len(value):
        start = _find_top_level_text_command(value, cursor)
        if start < 0:
            pieces.append(("math", value[cursor:]))
            break
        if start > cursor:
            pieces.append(("math", value[cursor:start]))
        open_brace = start + len(r"\text")
        end = _find_matching_brace(value, open_brace)
        if end < 0:
            pieces.append(("math", value[start:]))
            break
        pieces.append(("text", value[open_brace + 1:end]))
        cursor = end + 1

    rendered: List[str] = []
    for kind, piece in pieces:
        if kind == "text":
            rendered.append(escape_plain(piece.replace(r"\ ", " ")))
        else:
            math_piece = piece.strip()
            if math_piece:
                rendered.append(render_inline_math(math_piece))
    if not rendered:
        return escape_plain(value)
    return " ".join(part for part in rendered if part).replace("  ", " ")


def display_blocks(text: str) -> List[str]:
    return [m.group(1).strip() for m in re.finditer(r"\\\[(.*?)\\\]", text, flags=re.DOTALL)]


def normalize_align_environment(block: str) -> str:
    body = block.strip()
    env_match = re.search(r"\\begin\{(?:align\*?|aligned)\}(.*?)\\end\{(?:align\*?|aligned)\}", body, flags=re.DOTALL)
    if env_match:
        body = env_match.group(1).strip()
    body = re.sub(r"^=\s*", "", body)
    return body


def split_latex_rows(body: str) -> List[str]:
    parts = re.split(r"\\\\", body)
    return [part.strip() for part in parts if part.strip()]


def render_display_formula(block: str) -> str:
    body = normalize_align_environment(block)
    return "\\[\n\\begin{aligned}\n" + body + "\n\\end{aligned}\\]"


def render_align_block(block: str) -> str:
    body = normalize_align_environment(block)
    rows: List[str] = []
    for raw_row in split_latex_rows(body):
        row = raw_row.strip()
        if "&" in row:
            left, right = row.split("&", 1)
        else:
            left, right = row, ""
        left_html = latex_mixed_to_html(left)
        right_html = latex_mixed_to_html(right)
        rows.append(
            '<div class="math-align-row">'
            f'<span class="math-align-label">{left_html}</span>'
            f'<span class="math-align-eq">{right_html}</span>'
            '</div>'
        )
    return '<div class="math-align-block">' + "".join(rows) + '</div>'


def render_array_table(block: str) -> str:
    env = re.search(r"\\begin\{array\}\{[^}]*\}(.*?)\\end\{array\}", block, flags=re.DOTALL)
    body = env.group(1) if env else block
    rows_raw = split_latex_rows(body)
    rows: List[List[str]] = []
    for raw in rows_raw:
        cleaned = raw.replace(r"\hline", "").strip()
        if not cleaned:
            continue
        cells = [cell.strip() for cell in cleaned.split("&")]
        rows.append(cells)
    if not rows:
        return render_display_formula(block)
    html_rows: List[str] = []
    for r_index, row in enumerate(rows):
        tag = "th" if r_index == 0 else "td"
        html_rows.append("<tr>" + "".join(f"<{tag}>{latex_mixed_to_html(cell)}</{tag}>" for cell in row) + "</tr>")
    return '<div class="math-table-wrap"><table class="math-table">' + "".join(html_rows) + '</table></div>'


def render_units_block(block: str) -> str:
    env = re.search(r"\\begin\{array\}\{[^}]*\}(.*?)\\end\{array\}", block, flags=re.DOTALL)
    body = env.group(1) if env else block
    lines = [row.replace(r"\hline", "").strip() for row in split_latex_rows(body)]
    return '<div class="math-lines-block">' + "".join(
        f'<div class="math-line">{latex_mixed_to_html(line)}</div>' for line in lines if line
    ) + '</div>'


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.casefold()).strip("-")
    return slug or "section"


def split_explanation_sections(raw: str) -> List[Tuple[str, List[str]]]:
    lines = [line.rstrip() for line in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    sections: List[Tuple[str, List[str]]] = []
    current_heading: Optional[str] = None
    current_lines: List[str] = []

    def flush() -> None:
        nonlocal current_heading, current_lines
        if current_heading is not None:
            while current_lines and not current_lines[0].strip():
                current_lines.pop(0)
            while current_lines and not current_lines[-1].strip():
                current_lines.pop()
            sections.append((current_heading, current_lines))
        current_heading = None
        current_lines = []

    for line in lines:
        canonical = MAIN_HEADING_CANON.get(line.strip().casefold())
        if canonical:
            flush()
            current_heading = canonical
        else:
            if current_heading is None:
                current_heading = "Explanation:"
            current_lines.append(line)
    flush()
    return sections


def looks_like_subheading(line: str) -> bool:
    text = line.strip()
    return bool(text.endswith(":") and len(text) <= 110 and not text.startswith("\\") and not SVG_TOKEN_RE.search(text))


def render_theoretical(lines: Sequence[str], registry: SvgRegistry) -> str:
    clean = [line for line in lines if line.strip()]
    has_subheads = any(looks_like_subheading(line) for line in clean)
    if not has_subheads:
        body = render_lines(clean, "E", registry)
        return f'<div class="two-column-content" data-two-column-section="theoretical-explanation">{body}<br /></div>'

    groups: List[Tuple[Optional[str], List[str]]] = []
    heading: Optional[str] = None
    paragraph: List[str] = []
    for line in clean:
        if looks_like_subheading(line):
            if heading is not None or paragraph:
                groups.append((heading, paragraph))
            heading = line.strip()
            paragraph = []
        else:
            paragraph.append(line)
    if heading is not None or paragraph:
        groups.append((heading, paragraph))

    parts = ['<div class="two-column-content theoretical-explanation-flow" data-two-column-section="theoretical-explanation-group">']
    for subheading, paragraph_lines in groups:
        if subheading:
            parts.append(f'<b class="sol-subheading theoretical-flow-subheading">{escape_plain(subheading)}</b>')
        if paragraph_lines:
            slug = slugify(subheading[:-1] if subheading else "theoretical-explanation")
            paragraph_html = render_lines(paragraph_lines, "E", registry, joiner="<br />")
            parts.append(f'<div class="theoretical-flow-paragraph" data-theoretical-part="{slug}">{paragraph_html}</div>')
    parts.append('</div>')
    return "".join(parts)


def render_key_idea(lines: Sequence[str]) -> str:
    text = "\n".join(lines)
    blocks = display_blocks(text)
    if not blocks:
        return render_inline_text(text)
    parts: List[str] = [render_display_formula(blocks[0])]
    if len(blocks) >= 2:
        parts.append("<br />where<br />")
        parts.append(render_align_block(blocks[1]))
    for extra in blocks[2:]:
        parts.append("<br />" + render_display_formula(extra))
    return "".join(parts)


def render_explanation(raw: str, registry: SvgRegistry) -> str:
    out: List[str] = []
    for heading, lines in split_explanation_sections(raw):
        out.append(f'<b class="sol-subheading">{escape_plain(heading)}</b><br />')
        if heading == "Explanation:":
            continue
        if heading == "Concept Idea:":
            body = render_lines([line for line in lines if line.strip()], "E", registry)
            out.append(f'<div class="two-column-content" data-two-column-section="concept-idea">{body}<br /></div>')
        elif heading == "Illustration:":
            for line in lines:
                if line.strip():
                    out.append(render_line_with_svgs(line, "E", registry))
            out.append("<br /><br />")
        elif heading == "Key Idea:":
            out.append(render_key_idea(lines))
            out.append("<br />")
        elif heading == "Data Table:":
            blocks = display_blocks("\n".join(lines))
            out.append(render_array_table(blocks[0]) if blocks else render_lines(lines, "E", registry))
            out.append("<br />")
        elif heading == "Units Used:":
            blocks = display_blocks("\n".join(lines))
            out.append(render_units_block(blocks[0]) if blocks else render_lines(lines, "E", registry))
            out.append("<br />")
        elif heading == "Theoretical Explanation:":
            out.append(render_theoretical(lines, registry))
        elif heading == "Explanation Step by Step:":
            blocks = display_blocks("\n".join(lines))
            step_html = render_align_block(blocks[0]) if blocks else render_lines(lines, "E", registry)
            if step_html.startswith('<div class="math-align-block">'):
                step_html = step_html.replace(
                    '<div class="math-align-block">',
                    '<div class="math-align-block skillnox-step-by-step-block">',
                    1,
                )
            else:
                step_html = '<div class="skillnox-step-by-step-block">' + step_html + '</div>'
            out.append(step_html)
            out.append("<br />")
        else:
            section_name = slugify(heading[:-1])
            body = render_lines([line for line in lines if line.strip()], "E", registry)
            out.append(f'<div class="two-column-content" data-two-column-section="{section_name}">{body}<br /></div>')
    return "".join(out)


def render_solution(lines: Sequence[str], registry: SvgRegistry) -> str:
    body = render_lines([line for line in lines if line.strip()], "Q", registry)
    return f'<div class="two-column-content" data-two-column-section="solution">{body}</div>'


def render_option_content(option: OptionData, registry: SvgRegistry) -> str:
    return render_lines(option.lines, "Q", registry)


def _special_answer_attribute(spec: str) -> str:
    return html_lib.escape("/".join(_accepted_answer_values(spec)), quote=True)


def _render_special_marked_prompt(question: QuestionData, registry: SvgRegistry, slot_kind: str) -> str:
    slot_index = 0
    rendered_lines: List[str] = []
    for raw_line in question.question_lines:
        placeholders: List[Tuple[str, str]] = []

        def replace_marker(match: re.Match) -> str:
            nonlocal slot_index
            spec = match.group(1).strip()
            placeholder = f"SKILLNOX_SPECIAL_SLOT_{question.qkey}_{slot_index}_END"
            if slot_kind == "BLNK":
                accepted_attr = _special_answer_attribute(spec)
                replacement = (
                    f'<input id="blank-{question.qkey}-{slot_index + 1}" class="skillnox-blank-input" '
                    f'type="text" data-answers="{accepted_attr}" autocomplete="off" spellcheck="false" '
                    f'aria-label="Fill in the blank" placeholder="answer">'
                )
            else:
                answer_attr = html_lib.escape(re.sub(r"\s+", " ", spec.strip()), quote=True)
                replacement = (
                    f'<span class="skillnox-drop-zone" data-answer="{answer_attr}" tabindex="0">drop</span>'
                )
            placeholders.append((placeholder, replacement))
            slot_index += 1
            return placeholder

        marked = INLINE_ANSWER_MARK_RE.sub(replace_marker, raw_line)
        rendered = render_line_with_svgs(marked, "Q", registry)
        for placeholder, replacement in placeholders:
            rendered = rendered.replace(placeholder, replacement)
        rendered_lines.append(rendered)
    return "<br />".join(rendered_lines)


def _minimum_match_rotation(values: Sequence[str]) -> List[Tuple[int, str]]:
    indexed = list(enumerate(values, start=1))
    if len(indexed) <= 1:
        return indexed
    best = indexed
    best_hits = len(indexed) + 1
    for shift in range(1, len(indexed)):
        candidate = indexed[shift:] + indexed[:shift]
        hits = sum(
            1 for position, (_, value) in enumerate(candidate)
            if _normalise_special_answer(value) == _normalise_special_answer(values[position])
        )
        if hits < best_hits:
            best = candidate
            best_hits = hits
            if hits == 0:
                break
    return best


def _render_special_card_body(question: QuestionData, registry: SvgRegistry) -> str:
    qid = question.qkey
    qtype = normalize_question_type_code(question.qtype)
    if qtype == "TRFL":
        question_html = render_lines(question.question_lines, "Q", registry)
        if question.case_study_id:
            question_html = f'<span class="skillnox-case-part-marker" aria-hidden="true">({escape_plain(question.case_part)})</span> ' + question_html
        return f'''<div class="question-text">{question_html}</div>
  <div class="options-container">
    <div class="skillnox-trfl-list">
      <div class="option" data-val="TRUE" onclick="skillnoxSelectOption('{qid}', this, 'TRUE')"><span class="opt-text">True</span><span class="opt-icon" id="ic-{qid}-TRUE"></span></div>
      <div class="option" data-val="FALSE" onclick="skillnoxSelectOption('{qid}', this, 'FALSE')"><span class="opt-text">False</span><span class="opt-icon" id="ic-{qid}-FALSE"></span></div>
    </div>
    <div class="result-msg" id="resultMsg-{qid}"></div>
  </div>'''

    if qtype == "BLNK":
        question_html = _render_special_marked_prompt(question, registry, "BLNK")
        if question.case_study_id:
            question_html = f'<span class="skillnox-case-part-marker" aria-hidden="true">({escape_plain(question.case_part)})</span> ' + question_html
        return f'''<div class="question-text skillnox-interactive-question">{question_html}</div>
  <div class="options-container skillnox-interactive-container">
    <div class="result-msg" id="resultMsg-{qid}"></div>
  </div>'''

    if qtype == "DRWD":
        question_html = _render_special_marked_prompt(question, registry, "DRWD")
        if question.case_study_id:
            question_html = f'<span class="skillnox-case-part-marker" aria-hidden="true">({escape_plain(question.case_part)})</span> ' + question_html
        bank_rows: List[str] = []
        for token_index, answer in enumerate(question.answer_specs, start=1):
            word = re.sub(r"\s+", " ", answer.strip())
            word_attr = html_lib.escape(word, quote=True)
            bank_rows.append(
                f'<button type="button" class="skillnox-drag-word" draggable="true" '
                f'data-token-id="{qid}-token-{token_index}" data-word="{word_attr}">{escape_plain(word)}</button>'
            )
        return f'''<div class="question-text skillnox-interactive-question">{question_html}</div>
  <div class="options-container skillnox-interactive-container">
    <div class="skillnox-drag-note">Drag a word directly to a blank, or tap a word and then tap a blank. Filled blanks can be swapped. Repeated answers are separate tokens.</div>
    <div class="skillnox-drag-bank" aria-label="Drag word bank">{''.join(bank_rows)}</div>
    <div class="result-msg" id="resultMsg-{qid}"></div>
  </div>'''

    if qtype == "MTCH":
        initial = _minimum_match_rotation([right for _, right in question.match_pairs])
        rows: List[str] = []
        for (left, correct_answer), (token_source_index, initial_value) in zip(question.match_pairs, initial):
            left_html = render_line_with_svgs(left, "Q", registry)
            answer_attr = html_lib.escape(correct_answer, quote=True)
            initial_attr = html_lib.escape(initial_value, quote=True)
            token_id = f"{qid}-match-{token_source_index}"
            rows.append(
                f'<div class="skillnox-match-row" data-answer="{answer_attr}">'
                f'<div class="skillnox-match-left">{left_html}</div>'
                f'<div class="skillnox-match-slot" draggable="true" tabindex="0" '
                f'data-current-id="{token_id}" data-current="{initial_attr}" '
                f'data-initial-id="{token_id}" data-initial="{initial_attr}"><span>{escape_plain(initial_value)}</span></div></div>'
            )
        marker = ''
        if question.case_study_id:
            marker = f'<span class="skillnox-case-part-marker" aria-hidden="true">({escape_plain(question.case_part)})</span> '
        return f'''<div class="question-text">{marker}Match each item on the left with the correct answer on the right.</div>
  <div class="options-container skillnox-interactive-container">
    <div class="skillnox-drag-note">Drag one right-side answer onto another to swap them. On touch devices, tap one answer and then another. Duplicate answers are kept as separate instances.</div>
    <div class="skillnox-match-area">{''.join(rows)}</div>
    <div class="result-msg" id="resultMsg-{qid}"></div>
  </div>'''

    raise GenerationError(f"Unsupported special Question Type: {qtype}")


def _case_card_attributes(question: QuestionData) -> str:
    if not question.case_study_id:
        return ""
    return (
        f' data-case-study-group="{html_lib.escape(question.case_study_id, quote=True)}"'
        f' data-case-part="{html_lib.escape(question.case_part, quote=True)}"'
        f' data-case-display="{html_lib.escape(question.palette_number, quote=True)}"'
    )


def _question_number_text(question: QuestionData) -> str:
    return escape_plain(question.display_number)


def render_special_question_card(
    question: QuestionData,
    registry: SvgRegistry,
    question_attempt_max: int = 5,
) -> str:
    qid = question.qkey
    qtype = normalize_question_type_code(question.qtype)
    labels = {
        "TRFL": "TRFL · True / False",
        "BLNK": "BLNK · Fill in the Blank",
        "DRWD": "DRWD · Drag the Words",
        "MTCH": "MTCH · Match the Following",
    }
    body_html = _render_special_card_body(question, registry)
    sol_html = render_solution(question.solution_lines, registry)
    exp_html = render_explanation(question.explanation_raw, registry)
    sol_json = json.dumps(sol_html, ensure_ascii=True)
    exp_json = json.dumps(exp_html, ensure_ascii=True)
    qid_attr = escape_plain(question.qid)
    subtopic = escape_plain(question.subtopic)
    level = escape_plain(question.level)
    case_attrs = _case_card_attributes(question)

    return f'''<div class="glass-card" id="{qid}" QID="{qid_attr}" data-qtype="{qtype}"{case_attrs} data-has-sol="true" data-has-exp="true" data-is-checked="false" data-is-eye-clicked="false">
  <div class="q-meta" style="display:flex;flex-direction:column;align-items:flex-start;gap:4px;margin-bottom:22px;">
    <div class="q-number-box">Question {_question_number_text(question)}</div>
    <div class="q-subtopic-label" style="font-size:14px;font-weight:900;color:#fbbf24;margin-top:6px;letter-spacing:.02em;"><b>Subtopic: {subtopic}</b></div>
    <div class="skillnox-qtype-chip">{labels[qtype]}</div>
  </div>
  {body_html}
  <div id="questionStatusRow-{qid}" class="skillnox-question-status-row" style="display:none;">
    <span id="questionStatusPill-{qid}" class="skillnox-question-status-pill"></span>
    <button id="questionReviewBtn-{qid}" class="skillnox-question-review-btn" style="display:none;">Review Later</button>
  </div>
  <div class="actions">
    <div class="action-row-top">
      <button class="btn btn-submit" id="submitBtn-{qid}" onclick="skillnoxSubmitAnswer('{qid}')" disabled>Check</button>
      <span id="liveResult-{qid}" class="skillnox-live-result-badge" style="display:none;"></span>
      <span class="attempt-badge" id="attemptBadge-{qid}">Attempt: 0/{question_attempt_max}</span>
      <button class="btn btn-eye" id="eyeBtn-{qid}" style="display:none;" onclick="skillnoxRevealSolButtons('{qid}')">Show Answer</button>
      <div class="skillnox-question-level-badge skillnox-sl-level-badge" style="margin-left:auto;display:inline-flex;">{level}</div>
    </div>
    <div class="action-controls">
      <button class="btn btn-solution" id="btnSolution-{qid}" style="display:none;" onclick="skillnoxToggleSolution('{qid}')">&#10054; Solution</button>
      <button class="btn btn-explanation" id="btnExplanation-{qid}" style="display:none;" onclick="skillnoxToggleExplanation('{qid}')">&#10054; Explanation</button>
    </div>
    <div class="solution-panel" id="solutionPanel-{qid}"></div>
    <div class="solution-panel explanation-panel" id="explanationPanel-{qid}"></div>
  </div>
</div>
<script>
window.skillnoxState = window.skillnoxState || {{}};
window.skillnoxState['{qid}'] = {{
  correct: {json.dumps(question.correct)},
  solHTML: {sol_json},
  expHTML: {exp_json},
  selected: null,
  attemptCount: 0,
  isLocked: false,
  isChecked: false,
  isEyeClicked: false
}};
</script>'''


def render_question_card(
    question: QuestionData,
    registry: SvgRegistry,
    question_attempt_max: int = 5,
) -> str:
    if normalize_question_type_code(question.qtype) in SPECIAL_QUESTION_TYPES:
        return render_special_question_card(question, registry, question_attempt_max)

    qid = question.qkey
    question_html = render_lines(question.question_lines, "Q", registry)
    if question.case_study_id:
        question_html = (
            f'<span class="skillnox-case-part-marker" aria-hidden="true">({escape_plain(question.case_part)})</span> '
            + question_html
        )
    option_rows: List[str] = []
    for option in question.options:
        content = render_option_content(option, registry)
        frac_class = " option-frac" if "\\frac" in option.text else ""
        letter = option.letter
        option_rows.append(
            f'<div class="option{frac_class}" data-val="{letter}" onclick="skillnoxSelectOption(\'{qid}\', this, \'{letter}\')">'
            f'<span class="opt-text">{content}</span><span class="opt-icon" id="ic-{qid}-{letter}"></span></div>'
        )

    sol_html = render_solution(question.solution_lines, registry)
    exp_html = render_explanation(question.explanation_raw, registry)
    sol_json = json.dumps(sol_html, ensure_ascii=True)
    exp_json = json.dumps(exp_html, ensure_ascii=True)
    qid_attr = escape_plain(question.qid)
    subtopic = escape_plain(question.subtopic)
    level = escape_plain(question.level)
    qtype_attr = f' data-qtype="{html_lib.escape(normalize_question_type_code(question.qtype), quote=True)}"' if question.case_study_id else ""
    case_attrs = _case_card_attributes(question)

    return f'''<div class="glass-card" id="{qid}" QID="{qid_attr}"{qtype_attr}{case_attrs} data-has-sol="true" data-has-exp="true" data-is-checked="false" data-is-eye-clicked="false">
  <div class="q-meta" style="display: flex; flex-direction: column; align-items: flex-start; gap: 4px; margin-bottom: 22px;">
    <div class="q-number-box">Question {_question_number_text(question)}</div>
    <div class="q-subtopic-label" style="font-size: 14px; font-weight: 900; color: #fbbf24; margin-top: 6px; letter-spacing: 0.02em;"><b>Subtopic: {subtopic}</b></div>
  </div>
  <div class="question-text">{question_html}</div>
  <div class="options-container">
    <div class="options-list">
{chr(10).join(option_rows)}
    </div>
    <div class="result-msg" id="resultMsg-{qid}"></div>
  </div>
  <div id="questionStatusRow-{qid}" class="skillnox-question-status-row" style="display:none;">
    <span id="questionStatusPill-{qid}" class="skillnox-question-status-pill"></span>
    <button id="questionReviewBtn-{qid}" class="skillnox-question-review-btn" style="display:none;">
      Review Later
    </button>
  </div>
  <div class="actions">
    <div class="action-row-top">
      <button class="btn btn-submit" id="submitBtn-{qid}" onclick="skillnoxSubmitAnswer('{qid}')">Check</button>
      <span id="liveResult-{qid}" class="skillnox-live-result-badge" style="display:none;"></span>
      <span class="attempt-badge" id="attemptBadge-{qid}">Attempt: 0/{question_attempt_max}</span>
      <button class="btn btn-eye" id="eyeBtn-{qid}" style="display:none;" onclick="skillnoxRevealSolButtons('{qid}')">Show Answer</button>
      <div class="skillnox-question-level-badge skillnox-sl-level-badge" style="margin-left: auto; display: inline-flex;">{level}</div>
    </div>
    <div class="action-controls">
      <button class="btn btn-solution" id="btnSolution-{qid}" style="display:none;" onclick="skillnoxToggleSolution('{qid}')">&#10054; Solution</button>
      <button class="btn btn-explanation" id="btnExplanation-{qid}" style="display:none;" onclick="skillnoxToggleExplanation('{qid}')">&#10054; Explanation</button>
    </div>
    <div class="solution-panel" id="solutionPanel-{qid}"></div>
    <div class="solution-panel explanation-panel" id="explanationPanel-{qid}"></div>
  </div>
</div>
<script>
window.skillnoxState = window.skillnoxState || {{}};
window.skillnoxState['{qid}'] = {{
  correct: '{question.correct}',
  solHTML: {sol_json},
  expHTML: {exp_json},
  selected: null,
  attemptCount: 0,
  isLocked: false,
  isChecked: false,
  isEyeClicked: false
}};
</script>'''


def _case_parts_caption(parts: Sequence[str]) -> str:
    clean = [str(part).strip().lower() for part in parts if str(part).strip()]
    if not clean:
        return ""
    if len(clean) == 1:
        return f"({clean[0]})"
    return f"({clean[0]}-{clean[-1]})"


def render_case_study_passage(
    question: QuestionData,
    registry: SvgRegistry,
    bridge_after: str = "",
) -> str:
    case_id = html_lib.escape(question.case_study_id, quote=True)
    number = question.case_question_number or question.index
    caption = escape_plain(_case_parts_caption(question.case_parts))
    passage_html = render_lines(question.case_passage_lines, "Q", registry)
    subtopic = escape_plain(question.case_parent_subtopic or question.subtopic)
    bridge_class = " skillnox-case-study-bridge" if bridge_after else ""
    bridge_attr = f' data-case-study-bridge-after="{html_lib.escape(bridge_after, quote=True)}"' if bridge_after else ""
    id_attr = f' id="skillnoxCaseStudyPassage{number}"' if not bridge_after else ""
    return f'''<aside class="skillnox-case-study-passage{bridge_class}"{id_attr} data-case-study-copy="{case_id}"{bridge_attr}>
    <div class="skillnox-case-study-passage-head">
      <div class="skillnox-case-study-passage-title-wrap">
        <span class="q-number-box">Question {number}</span>
        <span class="skillnox-case-study-badge">CASE STUDY</span>
        <strong>Passage for Question {number} {caption}</strong>
      </div>
      <div class="q-subtopic-label" style="font-size:14px;font-weight:900;color:#fbbf24;margin-top:6px;letter-spacing:.02em;"><b>Subtopic: {subtopic}</b></div>
    </div>
    <div class="skillnox-case-study-passage-body">{passage_html}</div>
    <div class="skillnox-case-study-toggle-row"><button type="button" class="skillnox-case-study-toggle" aria-expanded="true" onclick="skillnoxToggleCaseStudyPassage('{case_id}', this)">Hide</button></div>
  </aside>'''


def render_question_region(
    questions: Sequence[QuestionData],
    registry: SvgRegistry,
    question_attempt_max: int = 5,
) -> str:
    out: List[str] = []
    i = 0
    while i < len(questions):
        question = questions[i]
        if not question.case_study_id:
            out.append(render_question_card(question, registry, question_attempt_max))
            i += 1
            continue

        case_id = question.case_study_id
        group: List[QuestionData] = []
        while i < len(questions) and questions[i].case_study_id == case_id:
            group.append(questions[i])
            i += 1
        first = group[0]
        number = first.case_question_number or first.index
        out.append(
            f'<section class="skillnox-case-study-group" id="skillnoxCaseStudy{number}" '
            f'data-case-study-id="{html_lib.escape(case_id, quote=True)}" '
            f'data-case-parent-qid="{html_lib.escape(first.case_parent_qid, quote=True)}" '
            f'data-case-question-number="{number}" aria-label="Case Study Question {number}">'
        )
        out.append(render_case_study_passage(first, registry))
        out.append('<div class="skillnox-case-study-questions">')
        for group_index, child in enumerate(group):
            block = render_question_card(child, registry, question_attempt_max)
            if group_index < len(group) - 1:
                state_marker = "\n<script>\nwindow.skillnoxState"
                card_html, sep, rest = block.partition(state_marker)
                if not sep:
                    raise GenerationError(f"Could not split state script for Case Study child {child.qkey}.")
                out.append(card_html)
                out.append(render_case_study_passage(first, registry, bridge_after=child.qkey))
                out.append(state_marker + rest)
            else:
                out.append(block)
        out.append('</div></section>')
    return "\n".join(out)


def patch_question_type_dashboard_labels(
    html_text: str,
    question_type_map: Dict[str, str],
) -> str:
    """Replace QTYPE short codes with configured full forms in Question Types analytics.

    The underlying question metadata keeps the stable short code (MCQ/BLNK/MTCH/etc.);
    this patch changes only the human-visible analytics label.
    """
    start = html_text.find('<div class="docx-dashboard question-types-dashboard"')
    if start < 0:
        return html_text
    end = html_text.find('<div class="glass-card"', start)
    if end < 0:
        end = len(html_text)
    section = html_text[start:end]
    for raw_code, raw_label in question_type_map.items():
        code = normalize_question_type_code(raw_code)
        label = str(raw_label or "").strip()
        if not code or not label:
            continue
        pattern = r'(<span>\s*)' + re.escape(code) + r'(\s*</span>)'
        section = re.sub(
            pattern,
            lambda m, text=escape_plain(label): m.group(1) + text + m.group(2),
            section,
            flags=re.IGNORECASE,
        )
    return html_text[:start] + section + html_text[end:]


def validate_question_type_dashboard_labels(
    html_text: str,
    questions: Sequence[QuestionData],
    question_type_map: Dict[str, str],
) -> None:
    """Verify each DOCX QTYPE used by the file has its mapped visible analytics label."""
    start = html_text.find('<div class="docx-dashboard question-types-dashboard"')
    if start < 0:
        raise GenerationError("Output validation failed; Question Types dashboard is missing.")
    end = html_text.find('<div class="glass-card"', start)
    if end < 0:
        end = len(html_text)
    section = html_text[start:end]
    for code in question_type_codes_from_questions(questions):
        label = str(question_type_map.get(code, "") or "").strip()
        if not label:
            raise GenerationError(f"Output validation failed; Question Type {code} has no full form.")
        if escape_plain(label) not in section:
            raise GenerationError(
                f"Output validation failed; Question Type label {label!r} for code {code} is missing from analytics."
            )


def render_dashboard(questions: Sequence[QuestionData], question_type_map: Dict[str, str]) -> str:
    subtopics: "OrderedDict[str, List[QuestionData]]" = OrderedDict()
    qtypes: "OrderedDict[str, int]" = OrderedDict()
    for question in questions:
        subtopics.setdefault(question.subtopic, []).append(question)
        qtypes[question.qtype] = qtypes.get(question.qtype, 0) + 1
    max_count = max((len(group) for group in subtopics.values()), default=1)

    groups: List[str] = []
    for subtopic, group in subtopics.items():
        first_q = group[0].qkey
        count = len(group)
        width = int(round((count / max_count) * 100))
        groups.append(f'''                            <div class="bar-graph-group">
                                <div class="bar-graph-label-row">
                                    
                                    <a href="#{first_q}" class="subtopic-scroll-link" title="Click to scroll to first question of {escape_plain(subtopic)}" onclick="scrollToElement('{first_q}'); return false;">
                                        {escape_plain(subtopic)}
                                    </a>
                                    <span class="bar-graph-count">{count} Ques</span>
                                </div>
                                <div class="bar-graph-row">
                                    <div class="bar-graph-track">
                                        <div class="bar-graph-fill" style="width: {width}%"></div>
                                    </div>
                                </div>
                            </div>''')

    qtype_rows = "\n".join(
        f'''                            <div class="docx-stat-row" data-question-type-code="{escape_plain(qtype)}">
                                <span>{escape_plain(question_type_map.get(normalize_question_type_code(qtype), qtype))}</span>
                                <strong style="color:var(--accent, #58a6ff)">{count}</strong>
                            </div>'''
        for qtype, count in qtypes.items()
    )
    total = len(questions)
    return f'''                <div class="docx-dashboard subtopics-dashboard">
                    <div class="docx-dash-section subtopics-section">
                        <h2>Subtopics Analytics</h2>
                        
{chr(10).join(groups)}
                    </div>
                </div>
                <div class="docx-dashboard question-types-dashboard" style="margin-top: 0;">
                    <div class="docx-dash-section question-types-section">
                        <h2>Question Types</h2>
                        
{qtype_rows}
                        <div class="docx-stat-row" style="margin-top:10px; border-top:1px solid var(--card-border, rgba(255,255,255,0.15)); padding-top:10px;">
                            <span><strong>Total Questions</strong></span>
                            <strong style="color:var(--accent, #58a6ff)">{total}</strong>
                        </div>
                    </div>
                </div>
    '''


# ---------------------------------------------------------------------------
# Golden template handling and exact-reference path
# ---------------------------------------------------------------------------

def decode_template_bytes() -> bytes:
    try:
        data = gzip.decompress(base64.b64decode(TEMPLATE_GZIP_B64.encode("ascii")))
    except Exception as exc:
        raise GenerationError(f"The bundled HTM template is corrupt: {exc}") from exc
    digest = hashlib.sha256(data).hexdigest()
    if digest != REFERENCE_HTML_SHA256:
        raise GenerationError(
            "The bundled HTM template failed its SHA-256 integrity check: " + digest
        )
    return data


def decode_template_text() -> str:
    return decode_template_bytes().decode("utf-8")


def replace_once(pattern: str, replacement: str, text: str, flags: int = 0, label: str = "pattern") -> str:
    updated, count = re.subn(pattern, lambda _: replacement, text, count=1, flags=flags)
    if count != 1:
        raise GenerationError(f"Could not patch {label}; expected one match, found {count}.")
    return updated


def format_question_ids(questions: Sequence[QuestionData]) -> str:
    return "[" + ",".join(json.dumps(q.qkey) for q in questions) + "]"


def format_correct_answers(questions: Sequence[QuestionData]) -> str:
    return "{" + ",".join(f'{json.dumps(q.qkey)}:{json.dumps(q.correct)}' for q in questions) + " }"


def format_question_meta(questions: Sequence[QuestionData]) -> str:
    rows: List[str] = []
    for q in questions:
        fields = [
            f'"questionNo": {json.dumps(q.palette_number)}',
            f'"subtopic": {json.dumps(q.subtopic)}',
            f'"qtype": {json.dumps(q.qtype)}',
            f'"level": {json.dumps(q.level)}',
            f'"pt": {json.dumps(q.pt)}',
            f'"featureQID": {json.dumps(q.qid)}',
        ]
        if q.case_study_id:
            fields.extend([
                f'"caseStudyId": {json.dumps(q.case_study_id)}',
                f'"casePart": {json.dumps(q.case_part)}',
                f'"caseParentQID": {json.dumps(q.case_parent_qid)}',
            ])
        rows.append(f'{json.dumps(q.qkey)}: {{ ' + ", ".join(fields) + ' }')
    return "{" + ",\n".join(rows) + " }"


def document_display_name(docx_path: Path) -> str:
    """Return the selected DOCX filename without its extension for visible titles."""
    value = docx_path.stem.strip()
    return value or "Generated Quiz"


def format_attempt_rules(rules: Sequence[TestAttemptRule]) -> str:
    payload = [rule.as_dict() for rule in rules]
    return json.dumps(payload, ensure_ascii=False, separators=(", ", ": "))


def format_qtype_config(qtype_map: Dict[str, str]) -> str:
    payload: "OrderedDict[str, str]" = OrderedDict()
    for raw_code, raw_label in qtype_map.items():
        code = normalize_qtype_code(raw_code)
        label = str(raw_label or "").strip()
        if code and label:
            payload[code] = label
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def patch_qtype_configuration(template: str, qtype_map: Dict[str, str]) -> str:
    """Inject the DOCX-used Qtype code/full-form map into the self-contained HTM."""
    block = "window.skillnoxQtypeConfig = " + format_qtype_config(qtype_map) + ";"
    return replace_once(
        r"window\.skillnoxQtypeConfig\s*=\s*\{.*?\};",
        block,
        template,
        flags=re.DOTALL,
        label="Qtype configuration",
    )


def patch_self_learning_filter_fresh_start(template: str) -> str:
    """Make every generated HTM open with All Levels + All Qtypes selected."""
    if 'localStorage.removeItem(STORAGE_KEY);' in template and 'var saved = null;' in template:
        return template
    pattern = (
        r'try\s*\{\s*'
        r'var\s+saved\s*=\s*JSON\.parse\('
        r'localStorage\.getItem\(STORAGE_KEY\)\s*\|\|\s*"null"\);'
    )
    replacement = (
        'try {\n'
        '    localStorage.removeItem(STORAGE_KEY);\n'
        '    var saved = null;'
    )
    return replace_once(
        pattern,
        replacement,
        template,
        flags=re.DOTALL,
        label="Self Learning filter fresh-start logic",
    )


def validate_self_learning_filter_fresh_start(html_text: str) -> None:
    """Verify that stale browser filters cannot hide questions on first load."""
    required = (
        'localStorage.removeItem(STORAGE_KEY);',
        'var saved = null;',
    )
    missing = [marker for marker in required if marker not in html_text]
    if missing:
        raise GenerationError(
            "Generated HTM did not receive the Self Learning filter reset patch: "
            + ", ".join(missing)
        )
    stale_loader = re.search(
        r'var\s+saved\s*=\s*JSON\.parse\(\s*localStorage\.getItem\(STORAGE_KEY\)',
        html_text,
    )
    if stale_loader:
        raise GenerationError(
            "Generated HTM can still restore a stale Self Learning filter from localStorage."
        )


def validate_qtype_configuration_patch(html_text: str, qtype_map: Dict[str, str]) -> None:
    match = re.search(
        r"window\.skillnoxQtypeConfig\s*=\s*(\{.*?\});",
        html_text,
        flags=re.DOTALL,
    )
    if not match:
        raise GenerationError("Generated HTM is missing the variable Qtype configuration block.")
    try:
        actual = json.loads(match.group(1))
    except json.JSONDecodeError as exc:
        raise GenerationError(f"Generated HTM contains invalid Qtype JSON: {exc}") from exc
    expected = OrderedDict((normalize_qtype_code(k), str(v).strip()) for k, v in qtype_map.items())
    if actual != expected:
        raise GenerationError("Qtype insertion failed; generated values differ from the GUI/CLI mappings.")


def patch_attempt_configuration(template: str, settings: AttemptSettings) -> str:
    """Patch only the variable attempt configuration block in the HTM shell."""
    block = (
        'window.skillnoxAttemptConfig = {\n'
        f'  questionAttemptMax: {settings.question_attempt_max},\n'
        f'  testAttemptMax: {settings.test_attempt_max},\n'
        f'  testAttemptRules: {format_attempt_rules(settings.test_attempt_rules)}\n'
        '};'
    )
    return replace_once(
        r"window\.skillnoxAttemptConfig\s*=\s*\{.*?\n\};",
        block,
        template,
        flags=re.DOTALL,
        label="attempt configuration",
    )


def validate_attempt_configuration_patch(html_text: str, settings: AttemptSettings) -> None:
    """Verify that GUI/CLI attempt values were inserted into the generated HTM."""
    match = re.search(
        r"window\.skillnoxAttemptConfig\s*=\s*\{\s*"
        r"questionAttemptMax\s*:\s*(\d+)\s*,\s*"
        r"testAttemptMax\s*:\s*(\d+)\s*,\s*"
        r"testAttemptRules\s*:\s*(\[.*?\])\s*\}",
        html_text,
        flags=re.DOTALL,
    )
    if not match:
        raise GenerationError("Generated HTM is missing the variable attempt configuration block.")

    question_max = int(match.group(1))
    test_max = int(match.group(2))
    try:
        rules = json.loads(match.group(3))
    except json.JSONDecodeError as exc:
        raise GenerationError(f"Generated HTM contains invalid Test Attempt Rules JSON: {exc}") from exc

    expected_rules = [rule.as_dict() for rule in settings.test_attempt_rules]
    if question_max != settings.question_attempt_max:
        raise GenerationError(
            f"Question Attempts Max insertion failed: expected {settings.question_attempt_max}, found {question_max}."
        )
    if test_max != settings.test_attempt_max:
        raise GenerationError(
            f"Default Test Attempts Max insertion failed: expected {settings.test_attempt_max}, found {test_max}."
        )
    if rules != expected_rules:
        raise GenerationError("Test Attempt Rules insertion failed; generated values differ from the GUI/CLI settings.")


def patch_initial_attempt_badges(
    html_text: str,
    question_attempt_max: int,
    expected_count: int,
) -> str:
    """Synchronize every initial Attempt badge with the selected GUI/CLI limit.

    The interactive runtime already reads ``window.skillnoxAttemptConfig``. This
    patch prevents the static first paint from showing the template default
    (for example 0/5) before the first answer changes it to the configured value
    (for example 1/7).
    """
    max_value = int(question_attempt_max)
    pattern = re.compile(
        r'(<span\b(?=[^>]*\bclass=["\'][^"\']*\battempt-badge\b[^"\']*["\'])'
        r'(?=[^>]*\bid=["\']attemptBadge-[^"\']+["\'])[^>]*>\s*Attempt:\s*0/)'
        r'\d+(\s*</span>)',
        flags=re.IGNORECASE,
    )
    updated, count = pattern.subn(
        lambda match: match.group(1) + str(max_value) + match.group(2),
        html_text,
    )
    if count != expected_count:
        raise GenerationError(
            "Initial Question Attempt badge synchronization failed: "
            f"expected {expected_count} badges, updated {count}."
        )
    return updated


def validate_initial_attempt_badges(
    html_text: str,
    question_attempt_max: int,
    expected_count: int,
) -> None:
    pattern = re.compile(
        r'<span\b(?=[^>]*\bclass=["\'][^"\']*\battempt-badge\b[^"\']*["\'])'
        r'(?=[^>]*\bid=["\']attemptBadge-[^"\']+["\'])[^>]*>\s*Attempt:\s*0/(\d+)\s*</span>',
        flags=re.IGNORECASE,
    )
    values = [int(value) for value in pattern.findall(html_text)]
    if len(values) != expected_count:
        raise GenerationError(
            "Generated HTM attempt-badge validation failed: "
            f"expected {expected_count} badges, found {len(values)}."
        )
    wrong = [value for value in values if value != int(question_attempt_max)]
    if wrong:
        raise GenerationError(
            "Generated HTM contains an initial Attempt badge that does not match "
            f"Question Attempts Max ({question_attempt_max})."
        )


def normalize_enabled_output_modes(enabled_modes: Optional[Sequence[str]]) -> Tuple[str, ...]:
    """Return a canonical non-empty mode tuple in Self/Test/CRM order."""
    if enabled_modes is None:
        return OUTPUT_MODE_ORDER
    requested = {str(mode or "").strip().lower() for mode in enabled_modes}
    unknown = sorted(requested - set(OUTPUT_MODE_ORDER))
    if unknown:
        raise GenerationError("Unknown HTM mode(s): " + ", ".join(unknown))
    modes = tuple(mode for mode in OUTPUT_MODE_ORDER if mode in requested)
    if not modes:
        raise GenerationError("Select at least one HTM mode: Self Learning, Test, or CRM.")
    return modes


def patch_enabled_output_modes(html_text: str, enabled_modes: Optional[Sequence[str]]) -> str:
    """Limit the HTM to the selected user-facing modes without changing core runtime blocks.

    When all three modes are selected the reference behavior is returned unchanged.
    For a subset, a small availability/boot controller is appended to the variable
    question-state script. Disabled core functions remain available internally so
    Test <-> CRM transitions can still use the existing state machine safely.
    """
    modes = normalize_enabled_output_modes(enabled_modes)
    if modes == OUTPUT_MODE_ORDER:
        return html_text

    boot_mode = "self" if "self" in modes else ("test" if "test" in modes else "crm")
    modes_json = json.dumps(list(modes), ensure_ascii=True)
    boot_json = json.dumps(boot_mode, ensure_ascii=True)

    controller_template = r'''
/* SKILLNOX GENERATED MODE AVAILABILITY START */
window.skillnoxGeneratedModes = __MODES_JSON__;
window.skillnoxGeneratedBootMode = __BOOT_JSON__;
(function () {
  "use strict";
  var allowed = window.skillnoxGeneratedModes.slice();
  var bootMode = window.skillnoxGeneratedBootMode;

  function allowedMode(mode) {
    return allowed.indexOf(String(mode || "").toLowerCase()) >= 0;
  }

  function installAvailabilityStyle() {
    var style = document.getElementById("skillnox-generated-mode-availability-style");
    if (!style) {
      style = document.createElement("style");
      style.id = "skillnox-generated-mode-availability-style";
      (document.head || document.documentElement).appendChild(style);
    }
    var selectors = [];
    if (!allowedMode("self")) {
      selectors.push('[data-snx-mode="self"]', '#modeOptSelf', 'button[onclick*="switchToSelfLearning()"]');
    }
    if (!allowedMode("test")) {
      selectors.push('[data-snx-mode="test"]', '#modeOptTest', 'button[onclick*="handleSkillnoxModeClick(\'test\')"]');
    }
    if (!allowedMode("crm")) selectors.push('[data-snx-mode="crm"]');
    style.textContent = selectors.length ? selectors.join(",") + "{display:none!important;}" : "";
    if (document.documentElement) {
      document.documentElement.setAttribute("data-skillnox-generated-modes", allowed.join(" "));
    }
  }

  function requestedModeFromTarget(target) {
    if (!target || !target.closest) return "";
    var direct = target.closest("[data-snx-mode]");
    if (direct) return String(direct.getAttribute("data-snx-mode") || "").toLowerCase();
    var legacy = target.closest("#modeOptSelf, #modeOptTest");
    if (legacy) return legacy.id === "modeOptTest" ? "test" : "self";
    var inline = target.closest('[onclick*="handleSkillnoxModeClick"]');
    if (inline) {
      var onclickText = String(inline.getAttribute("onclick") || "").toLowerCase();
      if (onclickText.indexOf("test") >= 0) return "test";
      if (onclickText.indexOf("crm") >= 0) return "crm";
      if (onclickText.indexOf("self") >= 0) return "self";
    }
    return "";
  }

  document.addEventListener("click", function (event) {
    var requested = requestedModeFromTarget(event.target);
    if (requested && !allowedMode(requested)) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  }, true);

  function ensureBootMode() {
    installAvailabilityStyle();
    if (bootMode === "self") {
      if (window.skillnoxClassroomMode === true || window.skillnoxMode === "test") {
        if (typeof window.switchToSelfLearning === "function") window.switchToSelfLearning();
      }
      return;
    }
    if (bootMode === "test") {
      var setupOpen = !!(document.body && document.body.classList.contains("skillnox-setup-open"));
      if (window.skillnoxMode !== "test" ||
          (window.skillnoxTestStarted !== true && window.skillnoxTestEnded !== true && !setupOpen)) {
        if (typeof window.showTestSetupPanel === "function") window.showTestSetupPanel();
      }
      return;
    }
    if (bootMode === "crm" && window.skillnoxClassroomMode !== true) {
      if (typeof window.enterSkillnoxClassroomMode === "function") window.enterSkillnoxClassroomMode();
    }
  }

  function start() {
    installAvailabilityStyle();
    window.setTimeout(ensureBootMode, 0);
    window.setTimeout(ensureBootMode, 180);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
/* SKILLNOX GENERATED MODE AVAILABILITY END */
'''
    controller = (
        controller_template
        .replace("__MODES_JSON__", modes_json)
        .replace("__BOOT_JSON__", boot_json)
    )

    pattern = re.compile(
        r'(<script\b[^>]*\bid=["\']skillnox-question-state-data-js["\'][^>]*>)(.*?)(</script>)',
        flags=re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(html_text)
    if not match:
        raise GenerationError(
            "Generated HTM is missing the variable question-state script needed for mode selection."
        )
    replacement = match.group(1) + match.group(2).rstrip() + controller + "\n" + match.group(3)
    return html_text[:match.start()] + replacement + html_text[match.end():]


def validate_enabled_output_modes(html_text: str, enabled_modes: Optional[Sequence[str]]) -> None:
    modes = normalize_enabled_output_modes(enabled_modes)
    marker = "SKILLNOX GENERATED MODE AVAILABILITY START"
    if modes == OUTPUT_MODE_ORDER:
        if marker in html_text:
            raise GenerationError(
                "Full three-mode output unexpectedly contains a restrictive mode controller."
            )
        return

    expected = "window.skillnoxGeneratedModes = " + json.dumps(list(modes), ensure_ascii=True) + ";"
    if marker not in html_text or expected not in html_text:
        raise GenerationError(
            "Generated HTM mode-selection controller is missing or does not match the GUI selection."
        )
    boot_mode = "self" if "self" in modes else ("test" if "test" in modes else "crm")
    expected_boot = "window.skillnoxGeneratedBootMode = " + json.dumps(boot_mode, ensure_ascii=True) + ";"
    if expected_boot not in html_text:
        raise GenerationError(
            "Generated HTM startup mode does not match the selected mode set."
        )


def patch_document_identity(template: str, docx_path: Path) -> str:
    """Patch only variable document identity fields; do not alter template logic."""
    display_name = document_display_name(docx_path)
    title = escape_plain(display_name) + " — Generated Quiz"
    content_id = re.sub(r"[^A-Za-z0-9_]+", "_", display_name).strip("_") + "_generated"

    template = replace_once(
        r"<title>.*?</title>",
        f"<title>{title}</title>",
        template,
        flags=re.DOTALL,
        label="document title",
    )
    template = replace_once(
        r'window\.skillnoxContentId\s*=\s*".*?";',
        f"window.skillnoxContentId = {json.dumps(content_id)};",
        template,
        flags=re.DOTALL,
        label="content ID",
    )
    # The latest approved shell removes the old main-banner/deck-title HTML.
    # Keep backward compatibility with older shells, but do not require that
    # retired visible-title node to exist.
    deck_title_pattern = r'<h1\s+class=["\']deck-title["\']>.*?</h1>'
    if re.search(deck_title_pattern, template, flags=re.DOTALL):
        template = replace_once(
            deck_title_pattern,
            '<h1 class="deck-title">' + escape_plain(display_name) + '</h1>',
            template,
            flags=re.DOTALL,
            label="visible document title",
        )
    return template


PACKED_EXPLANATIONS_SCRIPT_RE = re.compile(
    r'<script\s+id=["\']skillnox-packed-explanations["\'][^>]*>.*?</script>',
    re.IGNORECASE | re.DOTALL,
)
EXP_HTML_LITERAL_RE = re.compile(r'expHTML:\s*("(?:\\.|[^"\\])*")')


def _encode_skillnox_varint(value: int) -> bytes:
    if value < 0:
        raise ValueError("varint value must be nonnegative")
    out = bytearray()
    while True:
        byte = value & 0x7F
        value >>= 7
        if value:
            out.append(byte | 0x80)
        else:
            out.append(byte)
            break
    return bytes(out)


def _skillnox_lz_pack(data: bytes) -> bytes:
    if not data:
        return b""
    recent: Dict[bytes, List[int]] = {}
    out = bytearray()
    literal = bytearray()
    n = len(data)

    def flush_literal() -> None:
        nonlocal literal
        if not literal:
            return
        out.append(0)
        out.extend(_encode_skillnox_varint(len(literal)))
        out.extend(literal)
        literal = bytearray()

    def remember(pos: int) -> None:
        if pos + 4 > n:
            return
        key = data[pos:pos + 4]
        bucket = recent.setdefault(key, [])
        bucket.append(pos)
        if len(bucket) > 72:
            del bucket[:-72]

    i = 0
    while i < n:
        best_dist = 0
        best_len = 0
        if i + 4 <= n:
            key = data[i:i + 4]
            for pos in reversed(recent.get(key, [])):
                dist = i - pos
                if dist <= 0 or dist > 1_500_000:
                    continue
                length = 4
                max_len = min(n - i, 131071)
                while length < max_len:
                    source_index = i + length - dist
                    if source_index < 0 or data[source_index] != data[i + length]:
                        break
                    length += 1
                if length > best_len:
                    best_len = length
                    best_dist = dist
                    if length >= 1024:
                        break
        # Marker + distance varint + length varint needs a useful gain over a literal.
        if best_len >= 6:
            flush_literal()
            out.append(1)
            out.extend(_encode_skillnox_varint(best_dist))
            out.extend(_encode_skillnox_varint(best_len))
            end = i + best_len
            while i < end:
                remember(i)
                i += 1
        else:
            literal.append(data[i])
            remember(i)
            i += 1
            if len(literal) >= 16384:
                flush_literal()
    flush_literal()
    return bytes(out)


def _packed_explanations_script(explanations: Sequence[str]) -> str:
    payload = bytearray()
    for explanation in explanations:
        encoded = explanation.encode("utf-8")
        payload.extend(len(encoded).to_bytes(4, "little", signed=False))
        payload.extend(encoded)
    packed_b64 = base64.b64encode(_skillnox_lz_pack(bytes(payload))).decode("ascii")
    return '''<script id="skillnox-packed-explanations">(function(){"use strict";function v(a,p){var n=0,sh=0,b;do{b=a[p.i++];n|=(b&127)<<sh;sh+=7}while(b&128);return n}function u(s){var bin=atob(s),a=new Uint8Array(bin.length);for(var i=0;i<bin.length;i++)a[i]=bin.charCodeAt(i);var o=[],p={i:0};while(p.i<a.length){var t=a[p.i++];if(t===0){var l=v(a,p);for(var j=0;j<l;j++)o.push(a[p.i++])}else{var d=v(a,p),l=v(a,p),st=o.length-d;for(var j=0;j<l;j++)o.push(o[st+j])}}return new Uint8Array(o)}function txt(a){if(window.TextDecoder)return new TextDecoder("utf-8").decode(a);var r="",i=0,c,c2,c3,c4,cp;while(i<a.length){c=a[i++];if(c<128)r+=String.fromCharCode(c);else if(c<224){c2=a[i++];r+=String.fromCharCode(((c&31)<<6)|(c2&63))}else if(c<240){c2=a[i++];c3=a[i++];r+=String.fromCharCode(((c&15)<<12)|((c2&63)<<6)|(c3&63))}else{c2=a[i++];c3=a[i++];c4=a[i++];cp=((c&7)<<18)|((c2&63)<<12)|((c3&63)<<6)|(c4&63);cp-=65536;r+=String.fromCharCode(55296+(cp>>10),56320+(cp&1023))}}return r}var a=u("''' + packed_b64 + '''"),p=0,r=[];while(p<a.length){var l=a[p]|(a[p+1]<<8)|(a[p+2]<<16)|(a[p+3]<<24);p+=4;r.push(txt(a.subarray(p,p+l)));p+=l}window.__skillnoxPackedExps=r})();</script>'''


PACKED_EXPLANATIONS_SLOT = "<!--SKILLNOX_PACKED_EXPLANATIONS_SLOT-->"


def pack_generated_explanations(html_text: str, expected_count: int) -> str:
    # A regenerated document must never keep the bundled template's old payload.
    # Preserve the master shell's original packed-payload position so script
    # execution order remains identical to the approved HTM.
    html_text = PACKED_EXPLANATIONS_SCRIPT_RE.sub(PACKED_EXPLANATIONS_SLOT, html_text)
    matches = list(EXP_HTML_LITERAL_RE.finditer(html_text))
    if len(matches) != expected_count:
        raise GenerationError(
            f"Explanation packing found {len(matches)} literal expHTML values; expected {expected_count}."
        )
    explanations = [json.loads(match.group(1)) for match in matches]
    pieces: List[str] = []
    cursor = 0
    for index, match in enumerate(matches):
        pieces.append(html_text[cursor:match.start()])
        pieces.append(f"expHTML: window.__skillnoxPackedExps[{index}]")
        cursor = match.end()
    pieces.append(html_text[cursor:])
    html_text = "".join(pieces)
    script = _packed_explanations_script(explanations)
    if PACKED_EXPLANATIONS_SLOT in html_text:
        if html_text.count(PACKED_EXPLANATIONS_SLOT) != 1:
            raise GenerationError("Packed Explanation insertion slot must occur exactly once.")
        return html_text.replace(PACKED_EXPLANATIONS_SLOT, script, 1)
    first_card = html_text.find('<div class="glass-card" id="q1"')
    if first_card < 0:
        raise GenerationError("Cannot insert packed explanations because q1 is missing.")
    return html_text[:first_card] + script + "\n" + html_text[first_card:]



def consolidate_generated_state_scripts(region: str) -> str:
    """Merge generated per-question state scripts into the consolidated shell block.

    The consolidated approved HTM keeps all generated question state in one named
    script (skillnox-question-state-data-js). The renderer still produces one small
    state script per card because that keeps card rendering simple; this function
    removes those transient scripts, concatenates their JavaScript in question order,
    and appends one named state-data script to the rendered question region.
    """
    state_script_re = re.compile(
        r'<script>\s*(window\.skillnoxState\s*=\s*window\.skillnoxState\s*\|\|\s*\{\};\s*'
        r'window\.skillnoxState\[\'q\d+\'\]\s*=\s*\{.*?\};)\s*</script>',
        flags=re.DOTALL,
    )
    blocks: List[str] = []

    def take(match: re.Match[str]) -> str:
        blocks.append(match.group(1).strip())
        return ''

    cleaned = state_script_re.sub(take, region)
    expected = len(re.findall(r'<div class="glass-card" id="q\d+"', cleaned))
    if len(blocks) != expected:
        raise GenerationError(
            f"Question-state consolidation found {len(blocks)} state blocks; expected {expected}."
        )
    merged = '\n\n'.join(blocks)
    return cleaned.rstrip() + '\n<script id="skillnox-question-state-data-js">\n' + merged + '\n</script>\n'


def patch_case_study_runtime_generic(template: str) -> str:
    """Replace only sample-specific Case Study bindings with data-driven equivalents."""
    generic_bridge_css = (
        '.skillnox-case-study-questions > .glass-card:has(.explanation-panel.open) + '
        '.skillnox-case-study-bridge{display: block !important;}'
        '.skillnox-case-study-questions > .skillnox-case-study-bridge:'
        'has(~ .glass-card:has(.explanation-panel.open)){position: relative !important;top: auto !important;}'
    )

    if generic_bridge_css not in template:
        start_match = re.search(
            r'#q\d+:has\(#explanationPanel-q\d+\.open\)\s*\+\s*\.skillnox-case-study-bridge',
            template,
        )
        if not start_match:
            raise GenerationError('Could not locate the master Case Study bridge selector block.')
        css_start = start_match.start()
        css_end = template.find('body.skillnox-test-active .skillnox-case-study-bridge', css_start)
        if css_end < 0:
            raise GenerationError('Could not locate the end of the master Case Study bridge selector block.')
        template = template[:css_start] + generic_bridge_css + template[css_end:]

    generic_runtime = r"""(function(){
  "use strict";
  window.skillnoxCaseStudyData=window.skillnoxCaseStudyData||{};

  function cleanPart(value,index){
    value=String(value||'').trim().toLowerCase();
    if(value)return value;
    var romans=['i','ii','iii','iv','v','vi','vii','viii','ix','x'];
    return romans[index]||String(index+1);
  }

  function collectCaseData(){
    var result={};
    var groups=document.querySelectorAll('.skillnox-case-study-group[data-case-study-id]');
    for(var g=0;g<groups.length;g++){
      var group=groups[g];
      var caseId=String(group.getAttribute('data-case-study-id')||'').trim();
      if(!caseId)continue;
      var number=String(group.getAttribute('data-case-question-number')||'').trim();
      if(!number){
        var idMatch=String(group.id||'').match(/(\d+)$/);
        var caseMatch=caseId.match(/(\d+)$/);
        number=idMatch?idMatch[1]:(caseMatch?caseMatch[1]:String(g+1));
      }
      var parentQID=String(group.getAttribute('data-case-parent-qid')||'').trim();
      var cards=group.querySelectorAll('.skillnox-case-study-questions > .glass-card[data-case-study-group]');
      var childQids=[];
      var parts=[];
      for(var i=0;i<cards.length;i++){
        var card=cards[i];
        if(String(card.getAttribute('data-case-study-group')||'')!==caseId)continue;
        childQids.push(card.id);
        parts.push(cleanPart(card.getAttribute('data-case-part'),parts.length));
      }
      result[caseId]={
        caseId:caseId,
        group:group,
        questionNumber:number,
        parentQID:parentQID,
        childQids:childQids,
        parts:parts
      };
    }
    window.skillnoxCaseStudyData=result;
    return result;
  }
  window.collectSkillnoxCaseStudyData=collectCaseData;

  function applyCollapsedState(caseId,collapsed){
    var group=document.querySelector('[data-case-study-id="'+caseId+'"]');
    if(!group)return;
    var passages=group.querySelectorAll('.skillnox-case-study-passage');
    for(var i=0;i<passages.length;i++){
      passages[i].classList.toggle('is-collapsed',!!collapsed);
      var btn=passages[i].querySelector('.skillnox-case-study-toggle');
      if(btn){
        btn.textContent=collapsed?'Show':'Hide';
        btn.setAttribute('aria-expanded',collapsed?'false':'true');
      }
    }
  }

  window.skillnoxToggleCaseStudyPassage=function(caseId,button){
    var source=button&&button.closest?button.closest('.skillnox-case-study-passage'):null;
    var collapsed=source?!source.classList.contains('is-collapsed'):false;
    applyCollapsedState(caseId,collapsed);
    try{sessionStorage.setItem('skillnox.caseStudy.'+caseId+'.collapsed',collapsed?'1':'0');}catch(e){}
  };

  function restoreCollapsedState(data){
    Object.keys(data).forEach(function(caseId){
      var collapsed=false;
      try{collapsed=sessionStorage.getItem('skillnox.caseStudy.'+caseId+'.collapsed')==='1';}catch(e){}
      applyCollapsedState(caseId,collapsed);
    });
  }

  function syncCaseLabels(data){
    Object.keys(data).forEach(function(caseId){
      var item=data[caseId];
      item.childQids.forEach(function(qid,index){
        var card=document.getElementById(qid);
        if(!card)return;
        var part=item.parts[index]||cleanPart('',index);
        var box=card.querySelector('.q-number-box');
        if(box)box.textContent='Question '+item.questionNumber+' ('+part+')';
        card.setAttribute('data-case-display',item.questionNumber+'('+part+')');
        card.setAttribute('data-case-part',part);
        card.setAttribute('data-case-study-group',caseId);
      });
    });
  }

  function syncMeta(data){
    var meta=window.skillnoxQuestionMeta||{};
    Object.keys(data).forEach(function(caseId){
      var item=data[caseId];
      item.childQids.forEach(function(qid,index){
        if(!meta[qid])return;
        var part=item.parts[index]||cleanPart('',index);
        meta[qid].questionNo=item.questionNumber+'('+part+')';
        meta[qid].caseStudyId=caseId;
        meta[qid].casePart=part;
        meta[qid].caseParentQID=item.parentQID;
      });
    });
  }

  function init(){
    var data=collectCaseData();
    syncMeta(data);
    syncCaseLabels(data);
    restoreCollapsedState(data);
    if(typeof window.skillnoxApplyQuestionHeaderLayout==='function'){
      try{window.skillnoxApplyQuestionHeaderLayout();}catch(e){}
      syncCaseLabels(data);
    }
    if(typeof window.updateSkillnoxAnalyticsDashboard==='function'){
      try{window.updateSkillnoxAnalyticsDashboard();}catch(e){}
    }
  }

  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});
  else init();
})();"""

    if 'window.collectSkillnoxCaseStudyData=collectCaseData;' not in template:
        needle = 'window.skillnoxCaseStudyData.case30='
        pos = template.find(needle)
        if pos < 0:
            raise GenerationError('Could not locate the master sample-specific Case Study runtime.')
        runtime_start = template.rfind('(function(){', 0, pos)
        runtime_end = template.find('})();', pos)
        if runtime_start < 0 or runtime_end < 0:
            raise GenerationError('Could not isolate the master sample-specific Case Study runtime.')
        runtime_end += len('})();')
        template = template[:runtime_start] + generic_runtime + template[runtime_end:]

    forbidden = (
        'window.skillnoxCaseStudyData.case30=',
        "sessionStorage.getItem('skillnox.caseStudy.case30.collapsed')",
        "['q30','q31','q32','q33','q34'].forEach",
        '#skillnoxCaseStudy30:has(#explanationPanel-',
    )
    leaked = [marker for marker in forbidden if marker in template]
    if leaked:
        raise GenerationError('Sample-specific Case Study runtime remains after genericization: ' + ', '.join(leaked))
    return template


def _extract_runtime_blocks_in_order(text: str, tag: str) -> List[Tuple[str, str]]:
    blocks: List[Tuple[str, str]] = []
    pattern = re.compile(rf'<{tag}\b([^>]*)>(.*?)</{tag}>', flags=re.IGNORECASE | re.DOTALL)
    for match in pattern.finditer(text):
        attrs = match.group(1)
        id_match = re.search(r'\bid=["\']([^"\']+)["\']', attrs, flags=re.IGNORECASE)
        blocks.append((id_match.group(1) if id_match else '', match.group(2)))
    return blocks


def _normalize_fixed_runtime_for_contract(script_text: str) -> str:
    """Normalize only values that legitimately vary from one generated DOCX to another.

    This intentionally does *not* normalize function bodies, CRM behavior, event handlers,
    fixed constants, or UI logic. Any such runtime change still fails the contract.
    """
    replacements = (
        (r'window\.skillnoxContentId\s*=\s*["\'][^"\']*["\']\s*;', 'window.skillnoxContentId = __CONTENT_ID__;'),
        (r'window\.skillnoxAttemptConfig\s*=\s*\{.*?\n\};', 'window.skillnoxAttemptConfig = __ATTEMPT_CONFIG__;'),
        (r'window\.skillnoxQuestionIds\s*=\s*\[.*?\];', 'window.skillnoxQuestionIds = __QUESTION_IDS__;'),
        (r'window\.skillnoxCorrectAnswers\s*=\s*\{.*?\s\};', 'window.skillnoxCorrectAnswers = __CORRECT_ANSWERS__;'),
        (r'window\.skillnoxQuestionMeta\s*=\s*\{.*?\s\};', 'window.skillnoxQuestionMeta = __QUESTION_META__;'),
        (r'window\.skillnoxQtypeConfig\s*=\s*\{.*?\};', 'window.skillnoxQtypeConfig = __QTYPE_CONFIG__;'),
    )
    out = script_text
    for pattern, replacement in replacements:
        out = re.sub(pattern, replacement, out, count=1, flags=re.DOTALL)
    return out


def _normalize_application_runtime_for_contract(script_text: str) -> str:
    # Backward-compatible alias for older validation helpers.
    return _normalize_fixed_runtime_for_contract(script_text)


def validate_latest_runtime_exact(html_text: str, shell_reference_text: str) -> None:
    """Protect every fixed style/script block from latest-master regression."""
    expected_styles = _extract_runtime_blocks_in_order(shell_reference_text, 'style')
    actual_styles = _extract_runtime_blocks_in_order(html_text, 'style')
    if len(expected_styles) != len(actual_styles):
        raise GenerationError(
            f'Latest runtime contract failed: style block count changed ({len(expected_styles)} -> {len(actual_styles)}).'
        )
    if [x[0] for x in expected_styles] != [x[0] for x in actual_styles]:
        raise GenerationError('Latest runtime contract failed: style block order/IDs changed.')
    for (block_id, expected), (_, actual) in zip(expected_styles, actual_styles):
        if expected != actual:
            raise GenerationError(f'Latest runtime contract failed: style#{block_id or "(unnamed)"} changed.')

    expected_scripts = _extract_runtime_blocks_in_order(shell_reference_text, 'script')
    actual_scripts = _extract_runtime_blocks_in_order(html_text, 'script')
    if len(expected_scripts) != len(actual_scripts):
        raise GenerationError(
            f'Latest runtime contract failed: script block count changed ({len(expected_scripts)} -> {len(actual_scripts)}).'
        )
    if [x[0] for x in expected_scripts] != [x[0] for x in actual_scripts]:
        raise GenerationError('Latest runtime contract failed: script block order/IDs changed.')

    variable_ids = {'skillnox-packed-explanations', 'skillnox-question-state-data-js'}
    for (block_id, expected), (_, actual) in zip(expected_scripts, actual_scripts):
        if block_id in variable_ids:
            continue
        expected = _normalize_fixed_runtime_for_contract(expected)
        actual = _normalize_fixed_runtime_for_contract(actual)
        if expected != actual:
            raise GenerationError(f'Latest runtime contract failed: script#{block_id or "(unnamed)"} changed.')

    required = (
        'function buildSkillnoxCrmSequence()',
        'function activateSkillnoxCrmHierarchy()',
        'function restoreSkillnoxCrmPreviousView()',
        'skillnoxCrmSubtopicQuestions',
        'skillnox-crm-subtopic-qnum',
        'body.skillnox-crm-mode #skillnoxViewDockButton',
        'id="skillnox-crm-radial-controls-css"',
        'id="skillnox-crm-radial-controls-js"',
        'skillnoxCrmRadialShell',
        'skillnoxCrmRadialLauncher',
        'id="skillnox-crm-esvg-free-drag-css"',
        'id="skillnox-crm-esvg-free-drag-js"',
        "gesture.mode = 'pinch'",
        'data-skillnox-crm-esvg-pinching',
        'var SKILLNOX_CRM_ESVG_ZOOM_MIN = 0.50;',
        'var SKILLNOX_CRM_ESVG_ZOOM_MAX = 3.00;',
        "document.addEventListener('wheel', function (event)",
        'window.collectSkillnoxCaseStudyData=collectCaseData;',
    )
    missing = [marker for marker in required if marker not in html_text]
    if missing:
        raise GenerationError('Latest CRM/runtime contract is incomplete: ' + ', '.join(missing))

def patch_template_generic(
    template: str,
    docx_path: Path,
    questions: Sequence[QuestionData],
    registry: SvgRegistry,
    question_type_map: Dict[str, str],
    question_attempt_max: int = 5,
) -> str:
    # Replace only the sample-specific Case Study bindings with their data-driven equivalent.
    template = patch_case_study_runtime_generic(template)

    # Reserve the exact master-shell position of the packed Explanation payload.
    # The sample payload is replaced after the regenerated question states are built.
    template, packed_slot_count = PACKED_EXPLANATIONS_SCRIPT_RE.subn(PACKED_EXPLANATIONS_SLOT, template, count=1)
    if packed_slot_count != 1:
        raise GenerationError(
            f"Expected exactly one packed Explanation payload in the master shell; found {packed_slot_count}."
        )
    template = replace_once(
        r"window\.skillnoxQuestionIds\s*=\s*\[.*?\];",
        "window.skillnoxQuestionIds = " + format_question_ids(questions) + ";",
        template,
        flags=re.DOTALL,
        label="question IDs",
    )
    template = replace_once(
        r"window\.skillnoxCorrectAnswers\s*=\s*\{.*?\s\};",
        "window.skillnoxCorrectAnswers = " + format_correct_answers(questions) + ";",
        template,
        flags=re.DOTALL,
        label="correct-answer map",
    )
    template = replace_once(
        r"window\.skillnoxQuestionMeta\s*=\s*\{.*?\s\};",
        "window.skillnoxQuestionMeta = " + format_question_meta(questions) + ";",
        template,
        flags=re.DOTALL,
        label="question metadata",
    )

    first_card = template.find('<div class="glass-card" id="q1"')
    if first_card < 0:
        raise GenerationError("The template does not contain the first question card marker.")
    dashboard_start = template.rfind('<div class="docx-dashboard subtopics-dashboard"', 0, first_card)
    if dashboard_start < 0:
        dashboard_start = template.rfind('<div class="docx-dashboard"', 0, first_card)
    if dashboard_start < 0:
        raise GenerationError("The template dashboard marker was not found.")
    template = template[:dashboard_start] + render_dashboard(questions, question_type_map) + template[first_card:]

    first_card = template.find('<div class="glass-card" id="q1"')
    suffix_marker = template.find('\n  \n</div>\n<button class="floating-back-btn"', first_card)
    if suffix_marker < 0:
        suffix_marker = template.find('\n</div>\n<button class="floating-back-btn"', first_card)
    if suffix_marker < 0:
        raise GenerationError("The end of the question-card region was not found.")
    region = consolidate_generated_state_scripts(
        render_question_region(questions, registry, question_attempt_max)
    )
    template = template[:first_card] + region + template[suffix_marker:]
    return pack_generated_explanations(template, len(questions))


# ---------------------------------------------------------------------------
# Validation and orchestration
# ---------------------------------------------------------------------------

APPROVED_UI_IDS: Tuple[str, ...] = (
    "skillnoxControlDock",
    "skillnoxModeDockButton",
    "skillnoxModeDockPanel",
    "skillnoxViewDockButton",
    "skillnoxViewDockPanel",
    "skillnoxThemeBtn",
    "skillnoxThemeMenu",
    "skillnoxThemePicker",
    "skillnoxTestSetupPanel",
    "skillnoxTestSetupStage",
    "startConfiguredTestBtn",
    "skillnoxTestStickyBanner",
    "skillnoxTestResultPanel",
    "skillnoxQuestionPaletteOverlay",
    "skillnoxQuestionPaletteGrid",
    "skillnoxQuestionPaletteRailBtn",
    "skillnoxQuestionPaletteCloseRail",
    "skillnoxPaletteQuestionsTab",
    "skillnoxPaletteSubtopicsTab",
    "skillnoxSlideNavContainer",
    "skillnoxSlidePrevBtn",
    "skillnoxSlideNextBtn",
    "skillnoxSlideCounter",
    "skillnoxMainFilterToggle",
    "skillnoxMainFilterSummary",
    "skillnoxMainFilterReset",
    "skillnoxMainFilterProgress",
    "skillnoxQuestionParentDynamicButtons",
    "htm-zoom-container",
    "zoomToggleBtn",
    "zoomCtrlPanel",
    "floatingTopBtn",
)


def _named_element_ids(text: str, tag: str) -> set[str]:
    pattern = rf'<{tag}\b[^>]*\bid=["\']([^"\']+)["\']'
    return set(re.findall(pattern, text, flags=re.IGNORECASE))


def _named_js_functions(text: str) -> set[str]:
    return set(re.findall(r'\bfunction\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*\(', text))


def validate_approved_shell_contract(html_text: str, shell_reference_text: str) -> None:
    """Protect the approved consolidated shell from generic-render regressions.

    Question/dashboard contents are allowed to change, but the named CSS/JS runtime
    blocks, public JS functions and fixed UI objects must survive generation.
    This keeps the latest desktop/mobile/CRM HTM behavior as an enforced generator contract.
    """
    for tag in ("script", "style"):
        expected = _named_element_ids(shell_reference_text, tag)
        actual = _named_element_ids(html_text, tag)
        missing = sorted(expected - actual)
        if missing:
            raise GenerationError(
                f"Approved shell parity failed; missing named {tag} block(s): "
                + ", ".join(missing[:20])
            )

    # Packed Explanation and generated Question State payloads are variable data
    # blocks. Their internal helper functions are allowed to differ from the
    # bundled sample while the surrounding application runtime must stay intact.
    def shell_function_surface(text: str) -> set[str]:
        text = PACKED_EXPLANATIONS_SCRIPT_RE.sub("", text)
        text = re.sub(
            r'<script\b[^>]*\bid=["\']skillnox-question-state-data-js["\'][^>]*>.*?</script>',
            "",
            text,
            flags=re.IGNORECASE | re.DOTALL,
        )
        return _named_js_functions(text)

    expected_functions = shell_function_surface(shell_reference_text)
    actual_functions = shell_function_surface(html_text)
    missing_functions = sorted(expected_functions - actual_functions)
    if missing_functions:
        raise GenerationError(
            "Approved shell parity failed; missing JavaScript function(s): "
            + ", ".join(missing_functions[:25])
        )

    missing_ids = [
        element_id for element_id in APPROVED_UI_IDS
        if f'id="{element_id}"' not in html_text and f"id='{element_id}'" not in html_text
    ]
    if missing_ids:
        raise GenerationError(
            "Approved shell parity failed; missing fixed UI object(s): "
            + ", ".join(missing_ids)
        )

    required_dashboard_markers = (
        'class="docx-dashboard subtopics-dashboard"',
        'class="docx-dash-section subtopics-section"',
        'class="docx-dashboard question-types-dashboard"',
        'class="docx-dash-section question-types-section"',
        '>Subtopics Analytics<',
        '>Question Types<',
    )
    missing_dashboard = [m for m in required_dashboard_markers if m not in html_text]
    if missing_dashboard:
        raise GenerationError(
            "Approved dashboard structure was not preserved: "
            + ", ".join(missing_dashboard)
        )


def validate_output(
    html_text: str,
    questions: Sequence[QuestionData],
    report: GenerationReport,
    shell_reference_text: Optional[str] = None,
) -> None:
    required_runtime_markers = (
        'id="skillnox-startup-paint-guard-prepaint-js"',
        'id="skillnox-skeleton-prepaint-js"',
        'id="skillnox-consolidated-runtime-css"',
        'id="skillnox-application-runtime-js"',
        'skillnox-startup-pending',
        'skx-math-pending',
        'skillnoxReleaseStartupPaintGuard',
    )
    missing_runtime = [marker for marker in required_runtime_markers if marker not in html_text]
    if missing_runtime:
        raise GenerationError(
            "Output is missing approved consolidated runtime component(s): " + ", ".join(missing_runtime)
        )

    if shell_reference_text is not None:
        validate_approved_shell_contract(html_text, shell_reference_text)
        validate_latest_runtime_exact(html_text, shell_reference_text)

    if "base64" in html_text.casefold() or re.search(r"data\s*:\s*image/", html_text, flags=re.IGNORECASE):
        raise GenerationError(
            "Output would contain Base64/data-image content. Remove embedded raster data from the source SVG and try again."
        )

    unresolved = SVG_TOKEN_RE.findall(html_text)
    if unresolved:
        raise GenerationError("Unresolved SVG placeholder(s) remain in output: " + ", ".join(unresolved[:10]))

    for marker in (r"\frac{\)</span>", r"\\frac{\\)</span>"):
        if marker in html_text:
            raise GenerationError(
                "Output contains a split MathJax fraction near \\text{...}; nested TeX text must remain inside one math delimiter."
            )

    card_count = len(re.findall(r'<div class="glass-card" id="q\d+"', html_text))
    state_count = len(re.findall(r"window\.skillnoxState\['q\d+'\]\s*=", html_text))
    if card_count != len(questions):
        raise GenerationError(f"Output contains {card_count} question cards; expected {len(questions)}.")
    if state_count != len(questions):
        raise GenerationError(f"Output contains {state_count} question state blocks; expected {len(questions)}.")

    consolidated_singletons = (
        'skillnox-consolidated-runtime-css',
        'skillnox-application-runtime-js',
        'skillnox-question-state-data-js',
        'skillnox-packed-explanations',
        'skillnox-crm-qsvg-fit-css',
        'skillnox-crm-esvg-free-drag-css',
        'skillnox-crm-esvg-free-drag-js',
        'skillnox-crm-radial-controls-css',
        'skillnox-crm-radial-controls-js',
    )
    bad_singletons = [
        block_id for block_id in consolidated_singletons
        if html_text.count(f'id="{block_id}"') != 1
    ]
    if bad_singletons:
        raise GenerationError(
            "Consolidated runtime block(s) must occur exactly once: " + ", ".join(bad_singletons)
        )

    required_case_study_runtime = (
        'skillnox-case-study-passage',
        'skillnox-case-study-bridge',
        'skillnoxToggleCaseStudyPassage',
        'skillnox-case-study-palette-nav',
    )
    missing_case_runtime = [marker for marker in required_case_study_runtime if marker not in html_text]
    if missing_case_runtime:
        raise GenerationError(
            "Consolidated Case Study runtime is incomplete: " + ", ".join(missing_case_runtime)
        )

    required_crm_runtime = (
        'var SKILLNOX_CRM_ZOOM_STEPS = [1.20, 1.30, 1.40, 1.50, 1.60, 1.70, 1.80];',
        'var SKILLNOX_CRM_INITIAL_ZOOM = 1.50;',
        '--skillnox-crm-case-passage-height',
        'window.innerHeight * 0.30',
        'data-skillnox-crm-qsvg-fit',
        'var SKILLNOX_CRM_UNIT_SLIDE_ID = "skillnox-crm-unit-title";',
        'var SKILLNOX_CRM_SUBTOPIC_SLIDE_PREFIX = "skillnox-crm-subtopic-title:";',
        'function buildSkillnoxCrmSequence()',
        'function activateSkillnoxCrmHierarchy()',
        'function restoreSkillnoxCrmPreviousView()',
        'skillnox-crm-title-stage',
        'skillnoxCrmSubtopicQuestions',
        'skillnox-crm-subtopic-qnum',
        'body.skillnox-crm-mode #skillnoxViewDockButton',
        'id="skillnox-crm-radial-controls-css"',
        'id="skillnox-crm-radial-controls-js"',
        'skillnoxCrmRadialShell',
        'skillnoxCrmRadialLauncher',
        'id="skillnox-crm-esvg-free-drag-css"',
        'id="skillnox-crm-esvg-free-drag-js"',
        "gesture.mode = 'pinch'",
        'data-skillnox-crm-esvg-pinching',
        'var SKILLNOX_CRM_ESVG_ZOOM_MIN = 0.50;',
        'var SKILLNOX_CRM_ESVG_ZOOM_MAX = 3.00;',
        "document.addEventListener('wheel', function (event)",
        'window.collectSkillnoxCaseStudyData=collectCaseData;',
    )
    missing_crm = [marker for marker in required_crm_runtime if marker not in html_text]
    if missing_crm:
        raise GenerationError(
            "CRM runtime is incomplete: " + ", ".join(missing_crm)
        )

    obsolete_home_double_scale = (
        "fixedControls = document.querySelectorAll('.floating-back-btn, .floating-top-btn, .zoom-container"
    )
    if obsolete_home_double_scale in html_text:
        raise GenerationError(
            "CRM Home sizing regression detected: floating-top-btn is being counter-scaled twice."
        )
    packed_refs = re.findall(r'expHTML:\s*window\.__skillnoxPackedExps\[(\d+)\]', html_text)
    if len(packed_refs) != len(questions):
        raise GenerationError(f"Output contains {len(packed_refs)} packed explanation references; expected {len(questions)}.")

    legacy_markers = (
        'specialCorrect-',
        'skillnox-match-correct-line',
        'skillnox-interactive-answer-note',
        'Correct sequence:',
    )
    present_legacy = [marker for marker in legacy_markers if marker in html_text]
    if present_legacy:
        raise GenerationError("Superseded Question Type feedback code remains: " + ", ".join(present_legacy))

    for question in questions:
        qkey = question.qkey
        required = [
            f'id="{qkey}"',
            f'id="resultMsg-{qkey}"',
            f'id="questionStatusRow-{qkey}"',
            f'id="questionStatusPill-{qkey}"',
            f'id="questionReviewBtn-{qkey}"',
            f'id="submitBtn-{qkey}"',
            f'id="liveResult-{qkey}"',
            f'id="attemptBadge-{qkey}"',
            f'id="eyeBtn-{qkey}"',
            f'id="btnSolution-{qkey}"',
            f'id="btnExplanation-{qkey}"',
            f'id="solutionPanel-{qkey}"',
            f'id="explanationPanel-{qkey}"',
            f"window.skillnoxState['{qkey}']",
        ]
        for marker in required:
            if marker not in html_text:
                raise GenerationError(f"Output validation failed; missing marker: {marker}")

    groups: Dict[str, List[QuestionData]] = {}
    for question in questions:
        if question.case_study_id:
            groups.setdefault(question.case_study_id, []).append(question)
    for case_id, children in groups.items():
        first = children[0]
        number = first.case_question_number or first.index
        if f'data-case-study-id="{case_id}"' not in html_text:
            raise GenerationError(f"Case Study group {case_id} is missing.")
        if f'data-case-parent-qid="{first.case_parent_qid}"' not in html_text:
            raise GenerationError(f"Case Study parent QID for {case_id} is missing.")
        bridge_count = len(re.findall(rf'class="skillnox-case-study-passage skillnox-case-study-bridge"[^>]*data-case-study-copy="{re.escape(case_id)}"', html_text))
        expected_bridges = max(0, len(children) - 1)
        if bridge_count != expected_bridges:
            raise GenerationError(
                f"Case Study {case_id} contains {bridge_count} bridge passages; expected {expected_bridges}."
            )
        if f'data-case-question-number="{number}"' not in html_text:
            raise GenerationError(f"Case Study display number {number} is missing for {case_id}.")
        for child in children:
            meta_fragment = re.search(
                rf'{re.escape(json.dumps(child.qkey))}\s*:\s*\{{(.*?)\}}',
                html_text,
                flags=re.DOTALL,
            )
            if not meta_fragment:
                raise GenerationError(f"Case Study metadata for {child.qkey} is missing.")
            fragment = meta_fragment.group(1)
            for field_name, field_value in (
                ('caseStudyId', child.case_study_id),
                ('casePart', child.case_part),
                ('caseParentQID', child.case_parent_qid),
            ):
                expected = f'"{field_name}": {json.dumps(field_value)}'
                if expected not in fragment:
                    raise GenerationError(
                        f"Case Study metadata for {child.qkey} is missing {field_name}."
                    )

    forbidden_case_runtime = (
        'window.skillnoxCaseStudyData.case30=',
        "sessionStorage.getItem('skillnox.caseStudy.case30.collapsed')",
        "['q30','q31','q32','q33','q34'].forEach",
        '#skillnoxCaseStudy30:has(#explanationPanel-',
    )
    leaked = [marker for marker in forbidden_case_runtime if marker in html_text]
    if leaked:
        raise GenerationError(
            'Sample-specific Case Study runtime leaked into generated HTM: ' + ', '.join(leaked)
        )

    report.questions = len(questions)
    report.options = sum(len(q.options) for q in questions)
    report.qsvg_references = sum(
        len(SVG_TOKEN_RE.findall("\n".join(q.question_lines + [line for o in q.options for line in o.lines])))
        for q in questions
    )
    report.esvg_references = sum(len(SVG_TOKEN_RE.findall(q.explanation_raw)) for q in questions)


def default_output_path(docx_path: Path) -> Path:
    return docx_path.with_name(docx_path.stem + "_generated.htm")


def generate(
    docx_path: Path,
    output_path: Path,
    images_q: Path,
    images_e: Path,
    attempt_settings: Optional[AttemptSettings] = None,
    qtype_map: Optional[Dict[str, str]] = None,
    question_type_map: Optional[Dict[str, str]] = None,
    enabled_modes: Optional[Sequence[str]] = None,
) -> GenerationReport:
    tables = read_docx_tables(docx_path)
    table = select_question_table(tables)
    questions = parse_questions(table)
    effective_qtype_map = resolve_qtype_map_for_questions(questions, qtype_map)
    effective_question_type_map = resolve_question_type_map_for_questions(questions, question_type_map)
    effective_enabled_modes = normalize_enabled_output_modes(enabled_modes)
    q_folder = SvgFolder(images_q, "imagesQ")
    e_folder = SvgFolder(images_e, "imagesE")
    template_bytes = decode_template_bytes()
    template_text = template_bytes.decode("utf-8")
    template_text = patch_document_identity(template_text, docx_path)
    template_text = patch_qtype_configuration(template_text, effective_qtype_map)
    template_text = patch_self_learning_filter_fresh_start(template_text)
    effective_attempt_settings = attempt_settings or DEFAULT_ATTEMPT_SETTINGS
    template_text = patch_attempt_configuration(template_text, effective_attempt_settings)
    report = GenerationReport(
        exact_reference_mode=False,
        question_attempt_max=effective_attempt_settings.question_attempt_max,
        test_attempt_max=effective_attempt_settings.test_attempt_max,
        test_attempt_rule_count=len(effective_attempt_settings.test_attempt_rules),
        qtype_count=len(effective_qtype_map),
        qtype_codes=list(effective_qtype_map.keys()),
        question_type_count=len(effective_question_type_map),
        question_type_codes=list(effective_question_type_map.keys()),
        enabled_modes=list(effective_enabled_modes),
    )

    # One generation path only: every DOCX is rebuilt through the current
    # parser/render/runtime implementation. There is no old exact-reference path.
    registry = SvgRegistry(q_folder, e_folder)
    output_text = patch_template_generic(
        template_text,
        docx_path,
        questions,
        registry,
        effective_question_type_map,
        effective_attempt_settings.question_attempt_max,
    )

    output_text = patch_question_type_dashboard_labels(output_text, effective_question_type_map)
    output_text = patch_initial_attempt_badges(
        output_text,
        effective_attempt_settings.question_attempt_max,
        len(questions),
    )
    output_text = patch_enabled_output_modes(output_text, effective_enabled_modes)
    validate_qtype_configuration_patch(output_text, effective_qtype_map)
    validate_question_type_dashboard_labels(output_text, questions, effective_question_type_map)
    validate_self_learning_filter_fresh_start(output_text)
    validate_attempt_configuration_patch(output_text, effective_attempt_settings)
    validate_initial_attempt_badges(
        output_text,
        effective_attempt_settings.question_attempt_max,
        len(questions),
    )
    contract_reference_text = patch_case_study_runtime_generic(template_text)
    validate_output(output_text, questions, report, shell_reference_text=contract_reference_text)
    validate_enabled_output_modes(output_text, effective_enabled_modes)
    output_bytes = output_text.encode("utf-8")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    try:
        temp_path.write_bytes(output_bytes)
        os.replace(temp_path, output_path)
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass

    report.output_sha256 = hashlib.sha256(output_bytes).hexdigest()
    return report


def parse_cli_qtype(value: str) -> Tuple[str, str]:
    if "=" not in value:
        raise argparse.ArgumentTypeError("Use CODE=FULLFORM, for example G=Graphical")
    raw_code, raw_label = value.split("=", 1)
    code = normalize_qtype_code(raw_code)
    label = raw_label.strip()
    if not code:
        raise argparse.ArgumentTypeError("Qtype CODE cannot be blank")
    if not label:
        raise argparse.ArgumentTypeError("Qtype FULLFORM cannot be blank")
    return code, label


def qtype_map_from_args(args) -> "OrderedDict[str, str]":
    mapping = load_qtype_store()
    for code, label in (getattr(args, "qtype", None) or []):
        mapping[normalize_qtype_code(code)] = str(label).strip()
    return mapping


def parse_cli_question_type(value: str) -> Tuple[str, str]:
    if "=" not in value:
        raise argparse.ArgumentTypeError("Use CODE=FULLFORM, for example BLNK=Fill in the Blanks")
    raw_code, raw_label = value.split("=", 1)
    code = normalize_question_type_code(raw_code)
    label = raw_label.strip()
    if not code:
        raise argparse.ArgumentTypeError("Question Type CODE cannot be blank")
    if not label:
        raise argparse.ArgumentTypeError("Question Type FULLFORM cannot be blank")
    return code, label


def question_type_map_from_args(args) -> "OrderedDict[str, str]":
    mapping = load_question_type_store()
    for code, label in (getattr(args, "question_type", None) or []):
        mapping[normalize_question_type_code(code)] = str(label).strip()
    return mapping


def parse_cli_attempt_rule(value: str) -> TestAttemptRule:
    parts = [part.strip() for part in value.split(":")]
    if len(parts) != 3:
        raise argparse.ArgumentTypeError("Use FROM:TO:ATTEMPTS, for example 1:10:Unlimited")
    try:
        from_value = int(parts[0])
        to_value = int(parts[1])
    except ValueError as exc:
        raise argparse.ArgumentTypeError("FROM and TO must be integers") from exc
    if from_value < 1 or to_value < from_value:
        raise argparse.ArgumentTypeError("Rule range must satisfy 1 <= FROM <= TO")
    attempts_raw = parts[2]
    if attempts_raw.casefold() in {"unlimited", "0"}:
        attempts: AttemptValue = "unlimited"
    else:
        try:
            attempts = int(attempts_raw)
        except ValueError as exc:
            raise argparse.ArgumentTypeError("ATTEMPTS must be Unlimited, 0, or a positive integer") from exc
        if attempts <= 0:
            raise argparse.ArgumentTypeError("ATTEMPTS must be Unlimited, 0, or a positive integer")
    return TestAttemptRule(from_value, to_value, attempts)


def attempt_settings_from_args(args) -> AttemptSettings:
    question_max = int(getattr(args, "question_attempt_max", 5) or 5)
    if question_max < 1 or question_max > 99:
        raise GenerationError("--question-attempt-max must be from 1 to 99.")
    test_max = int(getattr(args, "test_attempt_max", 10) or 10)
    if test_max < 1 or test_max > 999:
        raise GenerationError("--test-attempt-max must be from 1 to 999.")
    supplied_rules = getattr(args, "test_attempt_rule", None)
    rules = tuple(supplied_rules) if supplied_rules else DEFAULT_ATTEMPT_SETTINGS.test_attempt_rules
    previous_to = 0
    for index, rule in enumerate(rules, start=1):
        if rule.from_value <= previous_to:
            raise GenerationError(f"Test attempt rule {index} overlaps or is out of order.")
        previous_to = rule.to_value
    return AttemptSettings(
        question_attempt_max=question_max,
        test_attempt_max=test_max,
        test_attempt_rules=rules,
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Generate the latest SkillNox no-Base64 HTM shell from a 15-column DOCX and sibling "
            "imagesQ/imagesE folders. With no DOCX argument, a graphical window opens and the "
            "finished HTM is launched automatically in Google Chrome. Qtype labels, QTYPE-column Question Type labels, and question/test limits are injected from the current GUI/CLI settings."
        )
    )
    parser.add_argument("docx", nargs="?", help="Input .docx file. Omit to open the graphical generator.")
    parser.add_argument("-o", "--output", help="Output .htm path. Default: <docx>_generated.htm")
    parser.add_argument("--images-q", help="Question SVG folder. Default: sibling imagesQ folder.")
    parser.add_argument("--images-e", help="Explanation SVG folder. Default: sibling imagesE folder.")
    parser.add_argument("--gui", action="store_true", help="Open the graphical generator window.")
    parser.add_argument("--open-chrome", action="store_true", help="In command-line mode, open the output in Chrome.")
    parser.add_argument("--no-open", action="store_true", help="Do not open the generated HTM automatically.")
    parser.add_argument(
        "--qtype",
        action="append",
        type=parse_cli_qtype,
        metavar="CODE=FULLFORM",
        help="Override/add a P/T Qtype full form for this run. Repeat for multiple Qtypes, e.g. --qtype G=Graphical.",
    )
    parser.add_argument(
        "--question-type",
        action="append",
        type=parse_cli_question_type,
        metavar="CODE=FULLFORM",
        help=(
            "Override/add a QTYPE-column Question Type full form for this run. "
            "Repeat for multiple codes, e.g. --question-type BLNK=Fill in the Blanks."
        ),
    )
    parser.add_argument(
        "--question-attempt-max",
        type=int,
        default=5,
        help="Maximum attempts for each question (1-99). Default: 5",
    )
    parser.add_argument(
        "--test-attempt-max",
        type=int,
        default=10,
        help="Fallback maximum test attempts when no range rule matches (1-999). Default: 10",
    )
    parser.add_argument(
        "--test-attempt-rule",
        action="append",
        type=parse_cli_attempt_rule,
        metavar="FROM:TO:ATTEMPTS",
        help=(
            "Test-attempt rule. Repeat for multiple ranges. "
            "ATTEMPTS may be Unlimited, 0, or a positive integer."
        ),
    )

    # Backward-compatible no-op flags. Old commands continue to work, but
    # neither flag enables SVG or golden-output matching.
    parser.add_argument("--strict-reference", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--allow-reference-differences", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--no-reinline-different-svg", action="store_true", help=argparse.SUPPRESS)

    parser.add_argument("--no-gui-message", action="store_true", help="Do not show a final Windows message box in CLI mode.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {APP_VERSION}")
    return parser


def run_cli(args) -> int:
    print(f"{APP_NAME} {APP_VERSION} — SVG comparison disabled")
    try:
        docx_path = Path(args.docx).expanduser().resolve()
        if not docx_path.is_file() or docx_path.suffix.lower() != ".docx":
            raise GenerationError(f"Select a valid .docx file: {docx_path}")

        base = docx_path.parent
        images_q = Path(args.images_q).expanduser().resolve() if args.images_q else locate_named_directory(base, "imagesQ")
        images_e = Path(args.images_e).expanduser().resolve() if args.images_e else locate_named_directory(base, "imagesE")
        if images_q is None or not images_q.is_dir():
            raise GenerationError(f"imagesQ folder not found beside the DOCX: {base / 'imagesQ'}")
        if images_e is None or not images_e.is_dir():
            raise GenerationError(f"imagesE folder not found beside the DOCX: {base / 'imagesE'}")

        output_path = Path(args.output).expanduser().resolve() if args.output else default_output_path(docx_path)
        attempt_settings = attempt_settings_from_args(args)
        qtype_map = qtype_map_from_args(args)
        question_type_map = question_type_map_from_args(args)
        report = generate(
            docx_path=docx_path,
            output_path=output_path,
            images_q=images_q,
            images_e=images_e,
            attempt_settings=attempt_settings,
            qtype_map=qtype_map,
            question_type_map=question_type_map,
        )

        mode_text = "Reference shell mode" if report.exact_reference_mode else "Generic DOCX mode"
        message = (
            f"Generated successfully.\n\n"
            f"Generator version: {APP_VERSION}\n"
            f"Output: {output_path}\n"
            f"Questions: {report.questions}\n"
            f"Options: {report.options}\n"
            f"QSVG references: {report.qsvg_references}\n"
            f"ESVG references: {report.esvg_references}\n"
            f"Qtypes (P/T): {', '.join(report.qtype_codes)}\n"
            f"Question Types (QTYPE): {', '.join(report.question_type_codes)}\n"
            f"Question Attempts Max: {attempt_settings.question_attempt_max}\n"
            f"Default Test Attempts Max: {attempt_settings.test_attempt_max}\n"
            f"Test Attempt Rules: {len(attempt_settings.test_attempt_rules)}\n"
            f"Mode: {mode_text}\n"
            f"SVG comparison: SKIPPED — current folder SVGs were inlined\n"
            f"SHA-256: {report.output_sha256}"
        )
        if report.warnings:
            message += "\n\nWarnings:\n- " + "\n- ".join(report.warnings[:12])

        if args.open_chrome and not args.no_open:
            _, browser_text = open_html_in_chrome(output_path)
            message += f"\n\n{browser_text}"

        print(message)
        if not args.no_gui_message:
            show_message("info", APP_NAME, message)
        return 0

    except GenerationError as exc:
        message = str(exc)
        print(f"ERROR: {message}", file=sys.stderr)
        if not getattr(args, "no_gui_message", False):
            show_message("error", APP_NAME, message)
        return 2
    except Exception as exc:
        details = traceback.format_exc()
        message = f"Unexpected error: {exc}\n\n{details}"
        print(message, file=sys.stderr)
        if not getattr(args, "no_gui_message", False):
            show_message("error", APP_NAME, message)
        return 3


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    # No DOCX path means the graphical workflow requested by the user.
    if args.gui or not args.docx:
        return launch_gui(args)
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
