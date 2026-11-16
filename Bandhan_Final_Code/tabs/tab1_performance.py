from __future__ import annotations

# ============================================================
# Combined Tab 1 + Tab 3 Automation
# Section 1: Performance Table Note
# Section 2: Fund Manager Count
# ============================================================

# QTabWidget is required for the Tab 5-style section tabs below.
from PyQt6.QtWidgets import QApplication, QLabel, QMainWindow, QTabWidget, QVBoxLayout, QWidget


# ============================================================
# SECTION 1 LOGIC: Performance Table Note (merged from tab1_performance_table_note.py)
# ============================================================

import argparse
import html
import json
import re
import sys
import traceback
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Callable, Iterable

import pandas as pd
from pptx import Presentation
from pptx.util import Pt

from PyQt6.QtCore import QObject, QRunnable, QThreadPool, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

tab1_APP_TITLE = "PPT Notes Automation - Performance Table Note"
tab1_STATE_FILE_NAME = "performance_table_note_state.json"
tab1_NOTE_FONT_NAME = "Ubuntu"
tab1_NOTE_FONT_SIZE_PT = 6
tab1_NOTES_SHEET_INDEX = 2
tab1_KNOWN_HTML_TAGS = {"sup", "sub", "b", "strong", "i", "em", "u", "br", "p", "div"}
tab1_PREFERRED_PPT_WORDS = ["fund", "manager", "performance"]


@dataclass
class tab1_PerformanceNote:
    scheme: str
    data_section: str
    note: str
    sequence: str = ""
    source: str = "Excel"
    use: bool = True


@dataclass
class tab1_ExclusionNote:
    note: str
    source: str = "Excel"
    use: bool = True


class tab1_WorkerSignals(QObject):
    started = pyqtSignal()
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)


class tab1_FunctionWorker(QRunnable):
    def __init__(self, func: Callable[..., Any], *args: Any, **kwargs: Any) -> None:
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.signals = tab1_WorkerSignals()

    @pyqtSlot()
    def run(self) -> None:
        self.signals.started.emit()
        try:
            if (
                hasattr(self.func, "__code__")
                and "progress_cb" in self.func.__code__.co_varnames
                and "progress_cb" not in self.kwargs
            ):
                self.kwargs["progress_cb"] = self.signals.progress.emit
            result = self.func(*self.args, **self.kwargs)
        except Exception:
            self.signals.error.emit(traceback.format_exc())
        else:
            self.signals.finished.emit(result)


# =========================================================
# Helpers
# =========================================================
def tab1_natural_key(text: str):
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", str(text))]


def tab1_normalize_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def tab1_normalize_sequence(value) -> str:
    text = tab1_normalize_text(value)
    if not text:
        return ""
    try:
        number = float(text)
        return str(int(number)) if number.is_integer() else str(number)
    except ValueError:
        return text


def tab1_sequence_sort_value(value) -> float:
    text = tab1_normalize_sequence(value)
    if not text:
        return 999999
    try:
        return float(text)
    except ValueError:
        return 999999


def tab1_excel_engine_for_file(path: str | Path) -> str:
    return "xlrd" if Path(path).suffix.lower() == ".xls" else "openpyxl"


def tab1_is_excel_file(path: Path) -> bool:
    return path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}


def tab1_get_files(folder: str | Path) -> list[Path]:
    folder = Path(folder)
    if not folder.exists() or not folder.is_dir():
        return []
    return [p for p in folder.iterdir() if p.is_file() and not p.name.startswith("~$")]


def tab1_clean_file_name_for_match(path: Path) -> str:
    text = path.stem.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tab1_find_pptx_file(folder: str | Path) -> Path | None:
    """
    Auto-detect source PPTX.

    Priority:
    1. PPTX containing exact normalized phrase: fund manager performance
    2. PPTX containing all words: fund, manager, performance
    3. First PPTX by natural sort
    """
    files = sorted([p for p in tab1_get_files(folder) if p.suffix.lower() == ".pptx"], key=lambda p: tab1_natural_key(p.name))
    files = [p for p in files if not p.name.startswith("~$")]
    if not files:
        return None

    exact_matches = [p for p in files if "fund manager performance" in tab1_clean_file_name_for_match(p)]
    if exact_matches:
        return exact_matches[0]

    word_matches = [p for p in files if all(word in tab1_clean_file_name_for_match(p) for word in tab1_PREFERRED_PPT_WORDS)]
    if word_matches:
        return word_matches[0]

    return files[0]


def tab1_find_client_input_excel(folder: str | Path) -> Path | None:
    excels = [p for p in tab1_get_files(folder) if tab1_is_excel_file(p)]
    preferred = [p for p in excels if p.name.lower().startswith("clientinputdata_")]
    if preferred:
        return sorted(preferred, key=lambda p: tab1_natural_key(p.name))[0]
    generic = [p for p in excels if "clientinputdata" in p.name.lower()]
    return sorted(generic, key=lambda p: tab1_natural_key(p.name))[0] if generic else None


def tab1_find_exclusion_excel(folder: str | Path) -> Path | None:
    matches = [p for p in tab1_get_files(folder) if tab1_is_excel_file(p) and "exclusion" in p.name.lower()]
    return sorted(matches, key=lambda p: tab1_natural_key(p.name))[0] if matches else None


# =========================================================
# Note text / HTML logic
# =========================================================
def tab1_decode_note_html(value) -> str:
    text = tab1_normalize_text(value)
    previous = None
    while previous != text:
        previous = text
        text = html.unescape(text)
    return text


def tab1_remove_only_known_html_tags(text) -> str:
    text = tab1_decode_note_html(text)
    for tag in tab1_KNOWN_HTML_TAGS:
        text = re.sub(rf"&lt;/?{tag}\b[^&gt;]*&gt;", " ", text, flags=re.I)
        text = re.sub(rf"</?{tag}\b[^>]*>", " ", text, flags=re.I)
    return text


def tab1_clean_match_text(value) -> str:
    text = tab1_remove_only_known_html_tags(value).lower()
    text = text.replace("&amp;", "and").replace("&", "and")
    text = text.replace("&lt;", " ").replace("&gt;", " ")
    text = text.replace("<", " ").replace(">", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tab1_compact_match_text(value) -> str:
    return tab1_clean_match_text(value).replace(" ", "")


def tab1_text_contains_value(haystack, needle) -> bool:
    cleaned_needle = tab1_clean_match_text(needle)
    if not cleaned_needle:
        return True
    cleaned_haystack = tab1_clean_match_text(haystack)
    return cleaned_needle in cleaned_haystack or tab1_compact_match_text(needle) in tab1_compact_match_text(haystack)


def tab1_note_is_excluded(note_text, exclusion_notes: Iterable[str]) -> bool:
    note_clean = tab1_clean_match_text(note_text)
    note_raw = tab1_decode_note_html(note_text)
    for exclusion in exclusion_notes:
        exclusion_clean = tab1_clean_match_text(exclusion)
        exclusion_raw = tab1_decode_note_html(exclusion)
        if exclusion_clean and (note_clean == exclusion_clean or (exclusion_raw and exclusion_raw in note_raw)):
            return True
    return False


def tab1_unique_notes_by_text(items: list[tab1_PerformanceNote]) -> list[tab1_PerformanceNote]:
    result = []
    seen = set()
    for item in items:
        key = tab1_clean_match_text(item.note)
        if key and key not in seen:
            seen.add(key)
            result.append(item)
    return result


def tab1_remove_exclusions_from_text(text, exclusion_notes: Iterable[str]) -> str:
    updated = tab1_decode_note_html(text)
    if not updated:
        return ""
    for exclusion in exclusion_notes:
        raw = tab1_decode_note_html(exclusion)
        clean = tab1_clean_match_text(raw)
        if not raw or not clean:
            continue
        updated = updated.replace(raw, "")
        if tab1_clean_match_text(updated) == clean:
            updated = ""
            continue
        chunks = re.split(r"(?<=[.!?])\s+|\n+", updated)
        updated = " ".join(chunk.strip() for chunk in chunks if tab1_clean_match_text(chunk) != clean and chunk.strip())
    updated = re.sub(r"\s+", " ", updated).strip()
    return re.sub(r"\s+([,.;:])", r"\1", updated)


class tab1_NoteHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=False)
        self.stack: list[str] = []
        self.segments: list[tuple[str, dict[str, bool]]] = []

    def current_style(self) -> dict[str, bool]:
        return {
            "sup": "sup" in self.stack,
            "sub": "sub" in self.stack,
            "bold": "b" in self.stack or "strong" in self.stack,
            "italic": "i" in self.stack or "em" in self.stack,
            "underline": "u" in self.stack,
        }

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in ["sup", "sub", "b", "strong", "i", "em", "u"]:
            self.stack.append(tag)
        elif tag in ["br", "p", "div"]:
            self.handle_data(" ")
        else:
            self.handle_data(f"<{tag}>")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ["sup", "sub", "b", "strong", "i", "em", "u"]:
            for i in range(len(self.stack) - 1, -1, -1):
                if self.stack[i] == tag:
                    del self.stack[i]
                    break
        elif tag not in ["br", "p", "div"]:
            self.handle_data(f"</{tag}>")

    def handle_data(self, data):
        if data:
            self.segments.append((html.unescape(data), self.current_style().copy()))

    def handle_entityref(self, name):
        self.handle_data(html.unescape(f"&{name};"))

    def handle_charref(self, name):
        self.handle_data(html.unescape(f"&#{name};"))


def tab1_parse_note_html(value):
    parser = tab1_NoteHTMLParser()
    parser.feed(tab1_decode_note_html(value))
    parser.close()
    return parser.segments


def tab1_apply_run_note_style(run, style: dict[str, bool]) -> None:
    run.font.name = tab1_NOTE_FONT_NAME
    run.font.size = Pt(tab1_NOTE_FONT_SIZE_PT)
    run.font.bold = bool(style.get("bold"))
    run.font.italic = bool(style.get("italic"))
    run.font.underline = bool(style.get("underline"))
    if style.get("sup") or style.get("sub"):
        run._r.get_or_add_rPr().set("baseline", "30000" if style.get("sup") else "-25000")


def tab1_set_note_shape_html_text(shape, note_html_text) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.clear()
    paragraph = tf.paragraphs[0]
    for text, style in tab1_parse_note_html(note_html_text):
        if text:
            run = paragraph.add_run()
            run.text = text
            tab1_apply_run_note_style(run, style)


# =========================================================
# PPT helpers
# =========================================================
def tab1_iterate_shapes(shapes):
    for shape in shapes:
        yield shape
        if hasattr(shape, "shapes"):
            yield from tab1_iterate_shapes(shape.shapes)


def tab1_get_shape_text(shape) -> str:
    parts = []
    if getattr(shape, "has_text_frame", False):
        parts.append(shape.text)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(x for x in parts if x)


def tab1_get_slide_text(slide) -> str:
    return "\n".join(text for text in (tab1_get_shape_text(shape) for shape in tab1_iterate_shapes(slide.shapes)) if text)


def tab1_find_note_shapes_on_slide(slide):
    return [shape for shape in tab1_iterate_shapes(slide.shapes) if getattr(shape, "name", "") == "Note"]

tab1_DEFAULT_PERFORMANCE_DISCLOSURE_NOTE = (
    "Performance based on NAV as on <NAVdate>. "
    "Past Performance may or may not be sustained in future. "
    "The performance details provided herein are of regular and direct plan growth option. "
    "Regular and Direct Plans have different expense structure. "
    "Direct Plan shall have a lower expense ratio excluding distribution expenses, commission expenses etc. "
    "Ms. Ritika Behera & Mr. Gaurav Satra has been designated as the dedicated fund manager for managing the overseas exposure "
    "in equity segment for all equity oriented schemes and hybrid schemes of Bandhan Mutual Fund (“the Fund”) with effect from "
    "October 07, 2023 and June 07, 2024 respectively."
)


def tab1_contains_performance_disclosure(text) -> bool:
    """Detect the standard performance disclosure even if NAV date changes."""
    cleaned = tab1_clean_match_text(text)
    return "performance based on nav as on" in cleaned and "past performance may or may not be sustained" in cleaned


def tab1_note_text_already_present(existing_text, new_note) -> bool:
    """Avoid appending duplicate notes. Uses exact clean match plus disclosure-prefix match."""
    if not tab1_clean_match_text(new_note):
        return True
    if tab1_contains_performance_disclosure(new_note) and tab1_contains_performance_disclosure(existing_text):
        return True
    return tab1_text_contains_value(existing_text, new_note)


def tab1_join_note_lines(*parts) -> str:
    """Join note parts on new lines while keeping existing PPT note text intact."""
    lines = []
    for part in parts:
        decoded = tab1_decode_note_html(part).strip()
        if decoded:
            lines.append(decoded)
    return "\n".join(lines)


def tab1_append_notes_to_existing_text(existing_text, notes_to_add: Iterable[str]) -> str:
    """Preserve existing Note textbox content and append only missing notes on new lines."""
    combined = tab1_decode_note_html(existing_text).strip()
    for note in notes_to_add:
        note = tab1_decode_note_html(note).strip()
        if not note:
            continue
        if not tab1_note_text_already_present(combined, note):
            combined = tab1_join_note_lines(combined, note)
    return combined


# =========================================================
# Excel helpers
# =========================================================
def tab1_get_notes_sheet_name(excel_path: str | Path, sheet_index: int = tab1_NOTES_SHEET_INDEX) -> str:
    xls = pd.ExcelFile(excel_path, engine=tab1_excel_engine_for_file(excel_path))
    return xls.sheet_names[sheet_index] if len(xls.sheet_names) > sheet_index else xls.sheet_names[-1]


def tab1_read_client_notes_from_excel(excel_path: str | Path, sheet_name: str | None = None) -> list[tab1_PerformanceNote]:
    engine = tab1_excel_engine_for_file(excel_path)
    sheet = sheet_name or tab1_get_notes_sheet_name(excel_path)
    df = pd.read_excel(excel_path, sheet_name=sheet, engine=engine, header=None)
    notes = []
    for _, row in df.iterrows():
        scheme = tab1_normalize_text(row.iloc[1] if len(row) > 1 else "")
        section = tab1_normalize_text(row.iloc[2] if len(row) > 2 else "")
        note = tab1_normalize_text(row.iloc[3] if len(row) > 3 else "")
        seq = tab1_normalize_sequence(row.iloc[4] if len(row) > 4 else "")
        if scheme and note and tab1_clean_match_text(scheme) not in {"scheme", "scheme name"}:
            notes.append(tab1_PerformanceNote(scheme, section, note, seq, "Excel", True))
    return notes


def tab1_read_exclusion_notes_from_excel(excel_path: str | Path | None) -> list[tab1_ExclusionNote]:
    if not excel_path or not Path(excel_path).exists():
        return []
    engine = tab1_excel_engine_for_file(excel_path)
    df = pd.read_excel(excel_path, engine=engine, header=None)
    result = []
    for value in df.values.flatten().tolist():
        text = tab1_normalize_text(value)
        if text and tab1_clean_match_text(text) not in {"exclusion", "exclusion notes", "note"}:
            result.append(tab1_ExclusionNote(text, "Excel", True))
    return result


# =========================================================
# Main processing
# =========================================================
def tab1_update_performance_table_notes(
    pptx_path: str | Path,
    output_path: str | Path,
    notes: list[tab1_PerformanceNote],
    exclusion_notes: Iterable[str] | None = None,
    progress_cb: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    def log(msg: str):
        if progress_cb:
            progress_cb(msg)

    pptx_path = Path(pptx_path)
    output_path = Path(output_path)
    exclusion_notes = list(exclusion_notes or [])
    if not pptx_path.exists():
        raise FileNotFoundError(f"Input PPTX not found: {pptx_path}")

    log("[Tab1] Progress: 0%")
    prs = Presentation(str(pptx_path))

    filtered = []
    excluded_count = 0
    for item in notes:
        if item.use and tab1_note_is_excluded(item.note, exclusion_notes):
            excluded_count += 1
        elif item.use:
            filtered.append(item)

    log(f"[Tab1] Notes selected: {len([n for n in notes if n.use])}")
    log(f"[Tab1] Notes excluded before adding: {excluded_count}")
    log(f"[Tab1] Notes to process: {len(filtered)}")
    log("[Tab1] Progress: 25%")

    # Do not clean/delete any existing text from PPT Note objects.
    # Existing/default PPT notes must stay exactly as text content and new notes are appended below.

    updated_slides = 0
    updated_shapes = 0
    missing_slides = []
    total_slides = max(1, len(prs.slides))
    log("[Tab1] Progress: 50%")

    for slide_index, slide in enumerate(prs.slides, start=1):
        slide_text = tab1_get_slide_text(slide)
        matching = []
        fallback = []
        schemes = []

        for item in filtered:
            scheme_matches = bool(item.scheme.strip()) and tab1_text_contains_value(slide_text, item.scheme)
            section_matches = tab1_text_contains_value(slide_text, item.data_section)
            if scheme_matches and section_matches:
                matching.append(item)
                schemes.append(item.scheme.strip())
            elif scheme_matches:
                fallback.append(item)

        used_fallback = False
        if not matching and fallback:
            matching = fallback
            schemes = [x.scheme.strip() for x in matching]
            used_fallback = True

        if matching:
            matching = tab1_unique_notes_by_text(sorted(matching, key=lambda x: tab1_sequence_sort_value(x.sequence)))
            note_shapes = tab1_find_note_shapes_on_slide(slide)
            if not note_shapes:
                missing_slides.append(slide_index)
                log(f"[Tab1] Slide {slide_index}: matched but no object named 'Note' found.")
            else:
                final_text = " ".join(x.note.strip() for x in matching if x.note.strip())
                final_text = tab1_remove_exclusions_from_text(final_text, exclusion_notes)

                notes_to_append = []
                if final_text.strip():
                    notes_to_append.append(final_text)

                # Always keep/add the standard performance disclosure.
                # If Excel already provides the same disclosure with an actual NAV date, this placeholder note is not duplicated.
                if not tab1_contains_performance_disclosure(final_text):
                    notes_to_append.append(tab1_DEFAULT_PERFORMANCE_DISCLOSURE_NOTE)

                for note_shape in note_shapes:
                    if getattr(note_shape, "has_text_frame", False):
                        existing_text = note_shape.text or ""
                        combined_text = tab1_append_notes_to_existing_text(existing_text, notes_to_append)
                        if combined_text != existing_text:
                            tab1_set_note_shape_html_text(note_shape, combined_text)
                            updated_shapes += 1
                updated_slides += 1
                log(f"[Tab1] Slide {slide_index}: appended note(s) without deleting existing/default PPT note text for scheme(s): {', '.join(sorted(set(schemes)))}")
                if used_fallback:
                    log(f"[Tab1] Slide {slide_index}: Data Section not found; used scheme-only fallback.")

        log(f"[Tab1] Progress: {50 + int((slide_index / total_slides) * 40)}%")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    log("[Tab1] Progress: 100%")
    log(f"[Tab1] Saved: {output_path}")
    return {
        "output_path": output_path,
        "notes_selected": len([n for n in notes if n.use]),
        "notes_excluded": excluded_count,
        "slides_updated": updated_slides,
        "note_shapes_updated": updated_shapes,
        "missing_note_shape_slides": missing_slides,
    }


def tab1_update_from_excel(pptx_path, client_excel_path, output_path, exclusion_excel_path=None, sheet_name=None, progress_cb=None):
    notes = tab1_read_client_notes_from_excel(client_excel_path, sheet_name)
    exclusions = [x.note for x in tab1_read_exclusion_notes_from_excel(exclusion_excel_path)]
    return tab1_update_performance_table_notes(pptx_path, output_path, notes, exclusions, progress_cb)


# =========================================================
# GUI
# =========================================================
class tab1_PerformanceTableNoteWindow(QWidget):
    NOTE_COL_USE, NOTE_COL_SCHEME, NOTE_COL_SECTION, NOTE_COL_NOTE, NOTE_COL_SEQUENCE, NOTE_COL_SOURCE = range(6)
    EXCL_COL_USE, EXCL_COL_NOTE, EXCL_COL_SOURCE = range(3)

    def __init__(self, state: Any | None = None):
        super().__init__()
        self.state = state
        self.thread_pool = QThreadPool.globalInstance()
        self.worker: tab1_FunctionWorker | None = None
        self.output_root: Path | None = getattr(state, "output_root", None) if state is not None else None
        self.default_output_folder = self._resolve_default_output_folder()
        self._build_ui()
        self._apply_default_output_folder()
        self._update_manual_note_controls_visibility()
        self._update_manual_exclusion_controls_visibility()
        self._set_progress_idle()


    def _app_run_folder(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parents[1]

    def _resolve_default_output_folder(self) -> Path:
        state_folder = getattr(self.state, "tab1_output_folder", None) if self.state is not None else None
        if state_folder:
            folder = Path(state_folder)
        else:
            output_root = getattr(self.state, "output_root", None) if self.state is not None else None
            folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "01_performance_table_note"
        folder.mkdir(parents=True, exist_ok=True)
        if self.state is not None:
            setattr(self.state, "tab1_output_folder", folder)
        return folder

    def _apply_default_output_folder(self) -> None:
        folder = self._resolve_default_output_folder()
        self.default_output_folder = folder
        if not self.output_folder_edit.text().strip():
            self.output_folder_edit.setText(str(folder))

    def _update_manual_note_controls_visibility(self) -> None:
        manual_enabled = self.chk_use_manual_notes.isChecked()
        self.btn_add_manual_note.setVisible(manual_enabled)

    def _update_manual_exclusion_controls_visibility(self) -> None:
        manual_enabled = self.chk_use_manual_exclusions.isChecked()
        self.btn_add_manual_exclusion.setVisible(manual_enabled)

    def _build_ui(self) -> None:
        self.setWindowTitle(tab1_APP_TITLE)
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, stretch=1)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(6, 6, 6, 6)
        left_layout.addWidget(QLabel("Process Logs"))
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        left_layout.addWidget(self.log_box, stretch=1)
        self.btn_clear_log = QPushButton("Clear Log")
        left_layout.addWidget(self.btn_clear_log)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(6, 6, 12, 6)
        layout.setSpacing(8)

        title = QLabel("Tab 1 - Performance Table Note")
        title.setStyleSheet("font-weight: bold; font-size: 15px;")
        desc = QLabel("Run this tab to create the performance-table notes-updated PPTX. Placeholders like <NAVdate> are preserved.")
        desc.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(desc)

        folder_group = QGroupBox("1. Folder Selection / Auto Detect")
        fg = QGridLayout(folder_group)
        self.source_folder_edit = QLineEdit()
        self.btn_browse_folder = QPushButton("Browse Folder...")
        self.btn_find_files = QPushButton("Find Files")
        fg.addWidget(QLabel("Source Folder:"), 0, 0)
        fg.addWidget(self.source_folder_edit, 0, 1)
        fg.addWidget(self.btn_browse_folder, 0, 2)
        fg.addWidget(self.btn_find_files, 0, 3)
        layout.addWidget(folder_group)

        input_group = QGroupBox("2. Tab 1 Input Files")
        ig = QGridLayout(input_group)
        self.pptx_edit = QLineEdit()
        self.client_excel_edit = QLineEdit()
        self.exclusion_excel_edit = QLineEdit()
        self.btn_browse_pptx = QPushButton("Browse...")
        self.btn_browse_client_excel = QPushButton("Browse...")
        self.btn_browse_exclusion_excel = QPushButton("Browse...")
        ig.addWidget(QLabel("Source PPTX:"), 0, 0)
        ig.addWidget(self.pptx_edit, 0, 1)
        ig.addWidget(self.btn_browse_pptx, 0, 2)
        ig.addWidget(QLabel("ClientInputData Excel:"), 1, 0)
        ig.addWidget(self.client_excel_edit, 1, 1)
        ig.addWidget(self.btn_browse_client_excel, 1, 2)
        ig.addWidget(QLabel("Exclusion_Notes Excel:"), 2, 0)
        ig.addWidget(self.exclusion_excel_edit, 2, 1)
        ig.addWidget(self.btn_browse_exclusion_excel, 2, 2)
        layout.addWidget(input_group)

        output_group = QGroupBox("3. Tab 1 Output")
        og = QGridLayout(output_group)
        self.output_folder_edit = QLineEdit()
        self.output_name_edit = QLineEdit()
        self.output_name_edit.setPlaceholderText("Example: Fund_Manager_Performance_Notes_Updated.pptx")
        og.addWidget(QLabel("Output Folder:"), 0, 0)
        og.addWidget(self.output_folder_edit, 0, 1)
        og.addWidget(QLabel("Output PPTX Name:"), 1, 0)
        og.addWidget(self.output_name_edit, 1, 1)
        layout.addWidget(output_group)

        notes_group = QGroupBox("4. Notes To Add")
        nl = QVBoxLayout(notes_group)
        nr = QHBoxLayout()
        self.chk_use_excel_notes = QCheckBox("Use Excel Notes")
        self.chk_use_excel_notes.setChecked(True)
        self.chk_use_manual_notes = QCheckBox("Use Manual Notes")
        nr.addWidget(self.chk_use_excel_notes)
        nr.addWidget(self.chk_use_manual_notes)
        nr.addStretch(1)
        nl.addLayout(nr)
        help_label = QLabel("Excel: B=Scheme, C=Data Section, D=Note, E=Sequence. Only real HTML tags are converted; placeholders like <NAVdate> stay unchanged.")
        help_label.setWordWrap(True)
        nl.addWidget(help_label)
        self.notes_table = QTableWidget(0, 6)
        self.notes_table.setHorizontalHeaderLabels(["Use?", "Scheme Name", "Data Section", "Note Text", "Sequence", "Source"])
        self.notes_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.notes_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.notes_table.setMinimumHeight(240)
        self._setup_notes_header()
        nl.addWidget(self.notes_table)
        nb = QHBoxLayout()
        self.btn_add_manual_note = QPushButton("Add Manual Note Row")
        self.btn_add_manual_note.setVisible(False)
        self.btn_remove_selected_notes = QPushButton("Remove Selected Note Row(s)")
        self.btn_remove_checked_notes = QPushButton("Remove Checked Note Row(s)")
        self.btn_check_all_notes = QPushButton("Check All Notes")
        self.btn_uncheck_all_notes = QPushButton("Uncheck All Notes")
        for button in [self.btn_add_manual_note, self.btn_remove_selected_notes, self.btn_remove_checked_notes, self.btn_check_all_notes, self.btn_uncheck_all_notes]:
            nb.addWidget(button)
        nb.addStretch(1)
        nl.addLayout(nb)
        layout.addWidget(notes_group)

        excl_group = QGroupBox("5. Exclusion Notes")
        el = QVBoxLayout(excl_group)
        er = QHBoxLayout()
        self.chk_use_excel_exclusions = QCheckBox("Use Excel Exclusions")
        self.chk_use_excel_exclusions.setChecked(True)
        self.chk_use_manual_exclusions = QCheckBox("Use Manual Exclusions")
        er.addWidget(self.chk_use_excel_exclusions)
        er.addWidget(self.chk_use_manual_exclusions)
        er.addStretch(1)
        el.addLayout(er)
        excl_help = QLabel("Exclusion Notes are never added; they remove matching notes from output.")
        excl_help.setWordWrap(True)
        el.addWidget(excl_help)
        self.exclusions_table = QTableWidget(0, 3)
        self.exclusions_table.setHorizontalHeaderLabels(["Use?", "Exclusion Note", "Source"])
        self.exclusions_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.exclusions_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.exclusions_table.setMinimumHeight(180)
        self._setup_excl_header()
        el.addWidget(self.exclusions_table)
        eb = QHBoxLayout()
        self.btn_add_manual_exclusion = QPushButton("Add Manual Exclusion Row")
        self.btn_add_manual_exclusion.setVisible(False)
        self.btn_remove_selected_exclusions = QPushButton("Remove Selected Exclusion Row(s)")
        self.btn_remove_checked_exclusions = QPushButton("Remove Checked Exclusion Row(s)")
        self.btn_check_all_exclusions = QPushButton("Check All Exclusions")
        self.btn_uncheck_all_exclusions = QPushButton("Uncheck All Exclusions")
        for button in [self.btn_add_manual_exclusion, self.btn_remove_selected_exclusions, self.btn_remove_checked_exclusions, self.btn_check_all_exclusions, self.btn_uncheck_all_exclusions]:
            eb.addWidget(button)
        eb.addStretch(1)
        el.addLayout(eb)
        layout.addWidget(excl_group)

        progress_group = QGroupBox("Progress")
        pl = QVBoxLayout(progress_group)
        self.progress_status_label = QLabel("Idle")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        pl.addWidget(self.progress_status_label)
        pl.addWidget(self.progress_bar)
        layout.addWidget(progress_group)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        self.btn_reload_excel = QPushButton("Reload Excel Data")
        self.btn_validation = QPushButton("Run Validation")
        self.btn_run = QPushButton("Run Performance Table Note")
        bottom.addWidget(self.btn_reload_excel)
        bottom.addWidget(self.btn_validation)
        bottom.addWidget(self.btn_run)
        layout.addLayout(bottom)

        scroll.setWidget(panel)
        splitter.addWidget(left)
        splitter.addWidget(scroll)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setSizes([340, 1080])

        self.btn_clear_log.clicked.connect(self.log_box.clear)
        self.btn_browse_folder.clicked.connect(self.browse_source_folder)
        self.btn_find_files.clicked.connect(self.find_files)
        self.btn_browse_pptx.clicked.connect(self.browse_pptx)
        self.btn_browse_client_excel.clicked.connect(self.browse_client_excel)
        self.btn_browse_exclusion_excel.clicked.connect(self.browse_exclusion_excel)
        self.btn_reload_excel.clicked.connect(self.reload_excel_data)
        self.btn_validation.clicked.connect(self.run_validation)
        self.btn_run.clicked.connect(self.start_process)
        self.btn_add_manual_note.clicked.connect(self.add_manual_note_row)
        self.chk_use_manual_notes.toggled.connect(self._update_manual_note_controls_visibility)
        self.btn_remove_selected_notes.clicked.connect(lambda: self.remove_selected_rows(self.notes_table))
        self.btn_remove_checked_notes.clicked.connect(lambda: self.remove_checked_rows(self.notes_table))
        self.btn_check_all_notes.clicked.connect(lambda: self.set_table_check_state(self.notes_table, True))
        self.btn_uncheck_all_notes.clicked.connect(lambda: self.set_table_check_state(self.notes_table, False))
        self.btn_add_manual_exclusion.clicked.connect(self.add_manual_exclusion_row)
        self.chk_use_manual_exclusions.toggled.connect(self._update_manual_exclusion_controls_visibility)
        self.btn_remove_selected_exclusions.clicked.connect(lambda: self.remove_selected_rows(self.exclusions_table))
        self.btn_remove_checked_exclusions.clicked.connect(lambda: self.remove_checked_rows(self.exclusions_table))
        self.btn_check_all_exclusions.clicked.connect(lambda: self.set_table_check_state(self.exclusions_table, True))
        self.btn_uncheck_all_exclusions.clicked.connect(lambda: self.set_table_check_state(self.exclusions_table, False))

    def _setup_notes_header(self):
        header = self.notes_table.horizontalHeader()
        for col in [self.NOTE_COL_SCHEME, self.NOTE_COL_SECTION, self.NOTE_COL_NOTE]:
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.Stretch)
        for col in [self.NOTE_COL_USE, self.NOTE_COL_SEQUENCE, self.NOTE_COL_SOURCE]:
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        self.notes_table.verticalHeader().setVisible(False)

    def _setup_excl_header(self):
        header = self.exclusions_table.horizontalHeader()
        header.setSectionResizeMode(self.EXCL_COL_NOTE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.EXCL_COL_USE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.EXCL_COL_SOURCE, QHeaderView.ResizeMode.ResizeToContents)
        self.exclusions_table.verticalHeader().setVisible(False)

    def log(self, text: str) -> None:
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def _set_progress_idle(self) -> None:
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_status_label.setText("Idle")

    def _set_progress_busy(self, text: str) -> None:
        self.progress_bar.setRange(0, 0)
        self.progress_status_label.setText(text)

    def _set_progress_value(self, value: int, text: str | None = None) -> None:
        clamped = max(0, min(100, int(value)))
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(clamped)
        self.progress_status_label.setText(text or f"{clamped}% complete")

    def on_progress(self, payload) -> None:
        match = re.search(r"(\d{1,3})\s*%", str(payload))
        if match:
            self._set_progress_value(int(match.group(1)))
        self.log(str(payload))

    def checkbox_item(self, checked=True):
        item = QTableWidgetItem()
        item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        return item

    def set_table_check_state(self, table, checked):
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            if item:
                item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def remove_selected_rows(self, table):
        rows = sorted({i.row() for i in table.selectionModel().selectedRows()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "No Selection", "Please select one or more rows.")
            return
        for row in rows:
            table.removeRow(row)

    def remove_checked_rows(self, table):
        for row in reversed(range(table.rowCount())):
            item = table.item(row, 0)
            if item and item.checkState() == Qt.CheckState.Checked:
                table.removeRow(row)

    def browse_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Source Folder", self.source_folder_edit.text() or str(Path.cwd()))
        if folder:
            self.source_folder_edit.setText(folder)
            if not self.output_folder_edit.text().strip():
                self.output_folder_edit.setText(str(self._resolve_default_output_folder()))
            self.find_files()

    def browse_pptx(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Source PPTX", self.source_folder_edit.text() or str(Path.cwd()), "PowerPoint Files (*.pptx)")
        if path:
            self.pptx_edit.setText(path)
            pptx = Path(path)
            if not self.source_folder_edit.text().strip():
                self.source_folder_edit.setText(str(pptx.parent))
            if not self.output_folder_edit.text().strip():
                self.output_folder_edit.setText(str(self._resolve_default_output_folder()))
            if not self.output_name_edit.text().strip():
                self.output_name_edit.setText(f"{pptx.stem}_performance_notes_updated.pptx")

    def browse_client_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select ClientInputData Excel", self.source_folder_edit.text() or str(Path.cwd()), "Excel Files (*.xlsx *.xlsm *.xls)")
        if path:
            self.client_excel_edit.setText(path)

    def browse_exclusion_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Exclusion_Notes Excel", self.source_folder_edit.text() or str(Path.cwd()), "Excel Files (*.xlsx *.xlsm *.xls)")
        if path:
            self.exclusion_excel_edit.setText(path)

    def find_files(self):
        folder_text = self.source_folder_edit.text().strip()
        if not folder_text or not Path(folder_text).is_dir():
            QMessageBox.warning(self, "Invalid Folder", "Please select a valid source folder.")
            return

        folder = Path(folder_text)
        pptx = tab1_find_pptx_file(folder)
        excel = tab1_find_client_input_excel(folder)
        exclusion = tab1_find_exclusion_excel(folder)

        if pptx:
            self.pptx_edit.setText(str(pptx))
            if not self.output_name_edit.text().strip():
                self.output_name_edit.setText(f"{pptx.stem}_performance_notes_updated.pptx")
        if excel:
            self.client_excel_edit.setText(str(excel))
        if exclusion:
            self.exclusion_excel_edit.setText(str(exclusion))
        if not self.output_folder_edit.text().strip():
            self.output_folder_edit.setText(str(folder))

        self.log("Auto-detect completed:")
        self.log("  Preferred PPT search: fund manager performance")
        self.log(f"  PPTX: {pptx or 'Not found'}")
        self.log(f"  ClientInputData Excel: {excel or 'Not found'}")
        self.log(f"  Exclusion Excel: {exclusion or 'Not found'}")

    def add_note_row(self, note: tab1_PerformanceNote):
        row = self.notes_table.rowCount()
        self.notes_table.insertRow(row)
        values = [None, note.scheme, note.data_section, note.note, note.sequence, note.source]
        self.notes_table.setItem(row, 0, self.checkbox_item(note.use))
        for col, val in enumerate(values[1:], start=1):
            self.notes_table.setItem(row, col, QTableWidgetItem(val))

    def add_manual_note_row(self):
        self.add_note_row(tab1_PerformanceNote("", "", "", "", "Manual", True))
        self.chk_use_manual_notes.setChecked(True)

    def add_exclusion_row(self, exclusion: tab1_ExclusionNote):
        row = self.exclusions_table.rowCount()
        self.exclusions_table.insertRow(row)
        self.exclusions_table.setItem(row, 0, self.checkbox_item(exclusion.use))
        self.exclusions_table.setItem(row, 1, QTableWidgetItem(exclusion.note))
        self.exclusions_table.setItem(row, 2, QTableWidgetItem(exclusion.source))

    def add_manual_exclusion_row(self):
        self.add_exclusion_row(tab1_ExclusionNote("", "Manual", True))
        self.chk_use_manual_exclusions.setChecked(True)

    def table_text(self, table, row, col):
        item = table.item(row, col)
        return item.text().strip() if item else ""

    def get_notes_data(self, include_unchecked=False):
        notes = []
        for row in range(self.notes_table.rowCount()):
            use_item = self.notes_table.item(row, 0)
            use = bool(use_item and use_item.checkState() == Qt.CheckState.Checked)
            if not include_unchecked and not use:
                continue
            source = self.table_text(self.notes_table, row, self.NOTE_COL_SOURCE) or "Excel"
            if source.lower() == "excel" and not self.chk_use_excel_notes.isChecked():
                continue
            if source.lower() == "manual" and not self.chk_use_manual_notes.isChecked():
                continue
            notes.append(tab1_PerformanceNote(
                self.table_text(self.notes_table, row, self.NOTE_COL_SCHEME),
                self.table_text(self.notes_table, row, self.NOTE_COL_SECTION),
                self.table_text(self.notes_table, row, self.NOTE_COL_NOTE),
                self.table_text(self.notes_table, row, self.NOTE_COL_SEQUENCE),
                source,
                use,
            ))
        return notes

    def get_exclusion_data(self, include_unchecked=False):
        exclusions = []
        for row in range(self.exclusions_table.rowCount()):
            use_item = self.exclusions_table.item(row, 0)
            use = bool(use_item and use_item.checkState() == Qt.CheckState.Checked)
            if not include_unchecked and not use:
                continue
            source = self.table_text(self.exclusions_table, row, self.EXCL_COL_SOURCE) or "Excel"
            if source.lower() == "excel" and not self.chk_use_excel_exclusions.isChecked():
                continue
            if source.lower() == "manual" and not self.chk_use_manual_exclusions.isChecked():
                continue
            exclusions.append(tab1_ExclusionNote(self.table_text(self.exclusions_table, row, self.EXCL_COL_NOTE), source, use))
        return exclusions

    def reload_excel_data(self):
        loaded_notes = 0
        loaded_exclusions = 0

        if self.chk_use_excel_notes.isChecked():
            excel = Path(self.client_excel_edit.text().strip())
            if not excel.exists():
                QMessageBox.warning(self, "Invalid Excel", "ClientInputData Excel does not exist.")
                return
            manual = [n for n in self.get_notes_data(True) if n.source.lower() == "manual"]
            self.notes_table.setRowCount(0)
            notes = tab1_read_client_notes_from_excel(excel)
            for note in notes + manual:
                self.add_note_row(note)
            loaded_notes = len(notes)

        if self.chk_use_excel_exclusions.isChecked():
            manual = [e for e in self.get_exclusion_data(True) if e.source.lower() == "manual"]
            self.exclusions_table.setRowCount(0)
            exclusion_text = self.exclusion_excel_edit.text().strip()
            exclusions = tab1_read_exclusion_notes_from_excel(exclusion_text) if exclusion_text else []
            for item in exclusions + manual:
                self.add_exclusion_row(item)
            loaded_exclusions = len(exclusions)

        self.log(f"Reloaded Excel data. Notes: {loaded_notes}, Exclusions: {loaded_exclusions}")
        QMessageBox.information(self, "Reload Complete", f"Notes loaded: {loaded_notes}\nExclusions loaded: {loaded_exclusions}")

    def validate_inputs(self):
        pptx_text = self.pptx_edit.text().strip()
        out_folder = self.output_folder_edit.text().strip()
        out_name = self.output_name_edit.text().strip()

        if not pptx_text or not Path(pptx_text).exists():
            QMessageBox.warning(self, "Invalid PPTX", "Please select a valid Source PPTX.")
            return None
        if not out_folder:
            QMessageBox.warning(self, "Invalid Output Folder", "Please select output folder.")
            return None
        Path(out_folder).mkdir(parents=True, exist_ok=True)
        if not out_name:
            QMessageBox.warning(self, "Invalid Output Name", "Please enter output PPTX name.")
            return None
        if not out_name.lower().endswith(".pptx"):
            out_name += ".pptx"

        if self.notes_table.rowCount() == 0 and self.chk_use_excel_notes.isChecked():
            self.reload_excel_data()

        notes = [n for n in self.get_notes_data() if n.use and n.scheme.strip() and n.note.strip()]
        if not notes:
            QMessageBox.warning(self, "No Notes", "No checked valid notes found.")
            return None
        exclusions = [e.note for e in self.get_exclusion_data() if e.use and e.note.strip()]
        return Path(pptx_text), Path(out_folder) / out_name, notes, exclusions

    def run_validation(self):
        valid = self.validate_inputs()
        if not valid:
            return
        pptx_path, output_path, notes, exclusions = valid
        try:
            prs = Presentation(str(pptx_path))
            scheme_matches = 0
            note_shape_matches = 0
            for slide in prs.slides:
                text = tab1_get_slide_text(slide)
                if any(tab1_text_contains_value(text, n.scheme) for n in notes):
                    scheme_matches += 1
                note_shape_matches += len(tab1_find_note_shapes_on_slide(slide))
            QMessageBox.information(
                self,
                "Validation Complete",
                f"PPTX:\n{pptx_path}\n\nOutput:\n{output_path}\n\nNotes selected: {len(notes)}\nExclusions selected: {len(exclusions)}\nSlides with scheme match: {scheme_matches}\nTotal Note objects found: {note_shape_matches}",
            )
        except Exception:
            error_text = traceback.format_exc()
            self.log(error_text)
            QMessageBox.critical(self, "Validation Failed", error_text)

    def start_process(self):
        valid = self.validate_inputs()
        if not valid:
            return
        pptx_path, output_path, notes, exclusions = valid
        self.log("=" * 70)
        self.log(f"[Tab1] Source PPTX: {pptx_path}")
        self.log(f"[Tab1] Output PPTX: {output_path}")
        self.btn_run.setEnabled(False)
        self.btn_reload_excel.setEnabled(False)
        self._set_progress_busy("Running Performance Table Note...")

        worker = tab1_FunctionWorker(tab1_update_performance_table_notes, pptx_path, output_path, notes, exclusions)
        self.worker = worker
        worker.signals.started.connect(lambda: self.log("[Tab1] Worker started"))
        worker.signals.progress.connect(self.on_progress)
        worker.signals.finished.connect(self.on_finished)
        worker.signals.error.connect(self.on_error)
        self.thread_pool.start(worker)

    def on_finished(self, result):
        self.worker = None
        self.btn_run.setEnabled(True)
        self.btn_reload_excel.setEnabled(True)
        self._set_progress_value(100, "Completed")
        output_path = result.get("output_path")
        if self.state is not None and output_path:
            setattr(self.state, "tab1_output_folder", Path(output_path).parent)
        self.log(f"[Tab1] Completed: {output_path}")
        self.log(f"[Tab1] Slides updated: {result.get('slides_updated')}")
        self.log(f"[Tab1] Note objects updated: {result.get('note_shapes_updated')}")
        QMessageBox.information(self, "Performance Table Note Complete", f"Performance Table Note completed successfully.\n\nSaved at:\n{output_path}")

    def on_error(self, error_text):
        self.worker = None
        self.btn_run.setEnabled(True)
        self.btn_reload_excel.setEnabled(True)
        self._set_progress_idle()
        self.progress_status_label.setText("Failed")
        self.log(error_text)
        QMessageBox.critical(self, "Performance Table Note Failed", error_text)

    # These methods are intentionally kept for compatibility with older direct usage,
    # but this tab no longer shows duplicate Choose Output Root / Save State / Load State buttons.
    def choose_output_root(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Output Root", str(self.output_root or Path.cwd()))
        if folder:
            self.output_root = Path(folder)
            if self.state is not None:
                setattr(self.state, "output_root", self.output_root)
            if not self.output_folder_edit.text().strip():
                self.output_folder_edit.setText(str(self._resolve_default_output_folder()))
            self.log(f"Output root selected: {self.output_root}")

    def save_state(self):
        data = {
            "output_root": str(self.output_root or ""),
            "source_folder": self.source_folder_edit.text(),
            "pptx": self.pptx_edit.text(),
            "client_excel": self.client_excel_edit.text(),
            "exclusion_excel": self.exclusion_excel_edit.text(),
            "output_folder": self.output_folder_edit.text(),
            "output_name": self.output_name_edit.text(),
            "use_excel_notes": self.chk_use_excel_notes.isChecked(),
            "use_manual_notes": self.chk_use_manual_notes.isChecked(),
            "use_excel_exclusions": self.chk_use_excel_exclusions.isChecked(),
            "use_manual_exclusions": self.chk_use_manual_exclusions.isChecked(),
        }
        path = (self.output_root or Path.cwd()) / tab1_STATE_FILE_NAME
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        self.log(f"State saved: {path}")

    def load_state(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load State", str(self.output_root or Path.cwd()), "JSON Files (*.json)")
        if not path:
            return
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        self.output_root = Path(data["output_root"]) if data.get("output_root") else None
        if self.state is not None:
            setattr(self.state, "output_root", self.output_root)
        self.source_folder_edit.setText(data.get("source_folder", ""))
        self.pptx_edit.setText(data.get("pptx", ""))
        self.client_excel_edit.setText(data.get("client_excel", ""))
        self.exclusion_excel_edit.setText(data.get("exclusion_excel", ""))
        self.output_folder_edit.setText(data.get("output_folder", ""))
        self.output_name_edit.setText(data.get("output_name", ""))
        self.chk_use_excel_notes.setChecked(bool(data.get("use_excel_notes", True)))
        self.chk_use_manual_notes.setChecked(bool(data.get("use_manual_notes", False)))
        self._update_manual_note_controls_visibility()
        self.chk_use_excel_exclusions.setChecked(bool(data.get("use_excel_exclusions", True)))
        self.chk_use_manual_exclusions.setChecked(bool(data.get("use_manual_exclusions", False)))
        self._update_manual_exclusion_controls_visibility()
        self.log(f"State loaded: {path}")


# Friendly alias for loaders that look for a generic widget class name.
tab1_Tab1Widget = tab1_PerformanceTableNoteWindow


def tab1_run_cli(args):
    tab1_update_from_excel(args.pptx, args.excel, args.output, args.exclusions or None, args.sheet or None, print)


def tab1_main():
    parser = argparse.ArgumentParser(description="Tab 1 - Performance Table Note updater")
    parser.add_argument("--pptx", default="")
    parser.add_argument("--excel", default="")
    parser.add_argument("--output", default="")
    parser.add_argument("--exclusions", default="")
    parser.add_argument("--sheet", default="")
    args = parser.parse_args()

    if args.pptx or args.excel or args.output:
        missing = [name for name, value in [("--pptx", args.pptx), ("--excel", args.excel), ("--output", args.output)] if not value]
        if missing:
            parser.error("CLI mode requires: " + ", ".join(missing))
        tab1_run_cli(args)
        return

    app = QApplication(sys.argv)
    window = tab1_PerformanceTableNoteWindow()
    window.resize(1435, 800)
    window.show()
    sys.exit(app.exec())


if False and __name__ == "__main__":
    tab1_main()


# ============================================================
# SECTION 2 LOGIC: Fund Manager Count PPT Checker (merged from tab3_fundmanager_count.py)
# ============================================================

import html
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pptx import Presentation
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


tab3_APP_TITLE = "Fund Manager Count PPT Checker - Easy Report"
tab3_DEFAULT_AMC = "Bandhan Mutual Fund"

tab3__ILLEGAL_EXCEL_CHARS = re.compile(
    r"[\x00-\x08\x0B-\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]"
)
tab3__FORMULA_START_CHARS = ("=", "+", "-", "@")
tab3__TITLE_PREFIX_RE = re.compile(r"^(mr|mrs|ms|dr)\.?\s*", flags=re.IGNORECASE)


# =========================================================
# TEXT HELPERS
# =========================================================

def tab3_now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def tab3_normalize_text(value) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    text = html.unescape(text)

    text = text.replace("\\<", "<").replace("\\>", ">")
    text = text.replace("\u00a0", " ")
    text = text.replace("–", "-")
    text = text.replace("—", "-")

    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"&lt;[^&gt;]+&gt;", " ", text)
    text = re.sub(r"[£†‡*§¥¢ß@^$]+", " ", text)

    return re.sub(r"\s+", " ", text).strip()


def tab3_clean_text(value) -> str:
    text = tab3_normalize_text(value).lower()
    text = text.replace("&amp;", "and")
    text = text.replace("&", "and")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tab3_manager_match_key(value) -> str:
    text = tab3_normalize_text(value)
    text = tab3__TITLE_PREFIX_RE.sub("", text).strip()
    return tab3_clean_text(text)


def tab3_scheme_match_key(value) -> str:
    return tab3_clean_text(value)


def tab3_split_hash_values(value) -> list[str]:
    text = tab3_normalize_text(value)

    if not text:
        return []

    return [tab3_normalize_text(part) for part in text.split("#") if tab3_normalize_text(part)]


def tab3_natural_key(text: str):
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", str(text))
    ]


def tab3_excel_engine(path: Path) -> str:
    return "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"


def tab3_scheme_word(count: int) -> str:
    return "scheme" if int(count) == 1 else "schemes"


def tab3_expected_note_text(manager: str, count: int, amc: str) -> str:
    return f"{manager} manages {count} {tab3_scheme_word(count)} of {amc}."


def tab3_amc_matches(excel_amc: str, ppt_amc: str) -> bool:
    excel_clean = tab3_clean_text(excel_amc)
    ppt_clean = tab3_clean_text(ppt_amc)

    if not excel_clean or not ppt_clean:
        return True

    return excel_clean == ppt_clean or excel_clean in ppt_clean or ppt_clean in excel_clean


# =========================================================
# SAFE EXCEL OUTPUT
# =========================================================

def tab3_safe_excel_value(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        text = value.strftime("%d-%b-%Y")
    elif isinstance(value, datetime):
        text = value.strftime("%d-%b-%Y")
    else:
        text = str(value)

    text = tab3__ILLEGAL_EXCEL_CHARS.sub("", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    if text.startswith(tab3__FORMULA_START_CHARS):
        text = "'" + text

    if len(text) > 30000:
        text = text[:30000] + " ...[trimmed]"

    return text


def tab3_safe_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    for col in out.columns:
        out[col] = out[col].map(tab3_safe_excel_value)

    return out


def tab3_style_workbook(path: Path):
    wb = load_workbook(path)

    dark_blue = "1F4E78"
    green = "C6EFCE"
    red = "FFC7CE"
    yellow = "FFF2CC"
    white = "FFFFFF"
    thin = Side(style="thin", color="D9D9D9")

    for ws in wb.worksheets:
        ws.title = "Easy Report"
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A2"

        for cell in ws[1]:
            cell.font = Font(bold=True, color=white, size=12)
            cell.fill = PatternFill("solid", fgColor=dark_blue)
            cell.alignment = Alignment(
                wrap_text=True,
                horizontal="center",
                vertical="center",
            )
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row in ws.iter_rows(min_row=2):
            result_value = str(row[1].value or "").upper() if len(row) > 1 else ""

            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

                if result_value == "OK":
                    cell.fill = PatternFill("solid", fgColor=green)
                elif result_value == "FIX":
                    cell.fill = PatternFill("solid", fgColor=red)
                else:
                    cell.fill = PatternFill("solid", fgColor=yellow)

        widths = {
            "A": 30,
            "B": 12,
            "C": 18,
            "D": 22,
            "E": 75,
            "F": 75,
            "G": 75,
            "H": 90,
            "I": 22,
        }

        for col_letter, width in widths.items():
            ws.column_dimensions[col_letter].width = width

        ws.auto_filter.ref = ws.dimensions

        for row_idx in range(1, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 28 if row_idx == 1 else 55

    wb.save(path)


def tab3_write_excel_report(output_path: Path, df: pd.DataFrame):
    output_path = Path(output_path)

    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(output_path.stem + "_temp.xlsx")

    if temp_path.exists():
        temp_path.unlink()

    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        tab3_safe_excel_df(df).to_excel(writer, index=False, sheet_name="Easy Report")

    tab3_style_workbook(temp_path)

    try:
        if output_path.exists():
            output_path.unlink()
        temp_path.replace(output_path)
    except PermissionError:
        raise PermissionError(
            f"Cannot replace report because it is open in Excel:\n"
            f"{output_path}\n\nClose the file and run again."
        )


# =========================================================
# EXCEL READING / COLUMN DETECTION
# =========================================================

def tab3_detect_header_row(raw_df: pd.DataFrame) -> int:
    keywords = [
        "scheme name",
        "fund name",
        "fund manager",
        "fund manager display",
        "amc",
        "mutual fund",
        "fund house",
    ]

    best_row = 0
    best_score = -1

    for row_idx in range(min(80, len(raw_df))):
        row_text = " ".join(tab3_clean_text(v) for v in raw_df.iloc[row_idx].tolist())
        score = sum(10 for kw in keywords if tab3_clean_text(kw) in row_text)

        if score > best_score:
            best_score = score
            best_row = row_idx

    return best_row


def tab3_dataframe_from_header(raw_df: pd.DataFrame, header_row: int) -> pd.DataFrame:
    headers = []
    seen = {}

    for idx, value in enumerate(raw_df.iloc[header_row].tolist()):
        header = tab3_normalize_text(value) or f"Column_{idx + 1}"

        if header in seen:
            seen[header] += 1
            header = f"{header}_{seen[header]}"
        else:
            seen[header] = 1

        headers.append(header)

    df = raw_df.iloc[header_row + 1:].copy()
    df.columns = headers
    return df.dropna(how="all").reset_index(drop=True)


def tab3_find_column(
    df: pd.DataFrame,
    exact_names: list[str],
    fallback_index: Optional[int] = None,
) -> Optional[str]:
    wanted = {tab3_clean_text(name) for name in exact_names}

    for col in df.columns:
        if tab3_clean_text(col) in wanted:
            return col

    for col in df.columns:
        col_clean = tab3_clean_text(col)

        if any(w in col_clean for w in wanted):
            return col

    if fallback_index is not None and fallback_index < len(df.columns):
        return df.columns[fallback_index]

    return None


def tab3_read_input_excel(path: Path, selected_sheet: str = "") -> tuple[pd.DataFrame, str, int, dict]:
    xls = pd.ExcelFile(path, engine=tab3_excel_engine(path))
    sheet = selected_sheet or xls.sheet_names[0]

    raw_df = pd.read_excel(
        path,
        engine=tab3_excel_engine(path),
        sheet_name=sheet,
        header=None,
    )

    if raw_df.empty:
        raise ValueError("Selected Excel sheet is empty.")

    header_row = tab3_detect_header_row(raw_df)
    df = tab3_dataframe_from_header(raw_df, header_row)

    detected = {
        "fund_name": tab3_find_column(df, ["Scheme Name", "Fund Name"], fallback_index=0),
        "fund_manager": tab3_find_column(df, ["Fund Manager"], fallback_index=1),
        "amc": tab3_find_column(
            df,
            ["AMC", "AMC Name", "Mutual Fund", "Fund House"],
            fallback_index=None,
        ),
    }

    return df, sheet, header_row, detected


def tab3_fundmanagerdetails_content_score(path: Path) -> tuple[int, Optional[str], Optional[int], dict]:
    try:
        xls = pd.ExcelFile(path, engine=tab3_excel_engine(path))
    except Exception:
        return -999, None, None, {}

    best_score = -999
    best_sheet = None
    best_header_row = None
    best_detected = {}

    for sheet in xls.sheet_names:
        try:
            raw_df = pd.read_excel(
                path,
                engine=tab3_excel_engine(path),
                sheet_name=sheet,
                header=None,
                nrows=120,
            )
        except Exception:
            continue

        if raw_df.empty:
            continue

        header_row = tab3_detect_header_row(raw_df)
        df = tab3_dataframe_from_header(raw_df, header_row)

        detected = {
            "fund_name": tab3_find_column(df, ["Scheme Name", "Fund Name"], fallback_index=0),
            "fund_manager": tab3_find_column(df, ["Fund Manager"], fallback_index=1),
            "amc": tab3_find_column(
                df,
                ["AMC", "AMC Name", "Mutual Fund", "Fund House"],
                fallback_index=None,
            ),
        }

        score = 0

        if detected["fund_name"]:
            score += 120
        if detected["fund_manager"]:
            score += 120
        if detected["amc"]:
            score += 20

        if not detected["fund_name"]:
            score -= 150
        if not detected["fund_manager"]:
            score -= 150

        name = path.name.lower()

        for bad_word in ["report", "simple_report", "check", "output", "exclusion", "notes"]:
            if bad_word in name:
                score -= 100

        for good_word in ["fundmanagerdetails", "fund_manager_details", "managerdetails", "manager"]:
            if good_word in name:
                score += 40

        if score > best_score:
            best_score = score
            best_sheet = sheet
            best_header_row = header_row
            best_detected = detected

    return best_score, best_sheet, best_header_row, best_detected


# =========================================================
# FILE AUTO DETECTION
# =========================================================

def tab3_auto_detect_files_from_folder(folder: Path) -> dict:
    files = [p for p in folder.iterdir() if p.is_file() and not p.name.startswith("~$")]
    excel_files = [p for p in files if p.suffix.lower() in [".xlsx", ".xlsm", ".xls"]]
    pptx_files = [p for p in files if p.suffix.lower() == ".pptx"]

    excel_candidates = []

    for excel_file in excel_files:
        score, sheet, header_row, detected = tab3_fundmanagerdetails_content_score(excel_file)
        excel_candidates.append(
            {
                "path": excel_file,
                "score": score,
                "sheet": sheet,
                "header_row": header_row,
                "detected": detected,
            }
        )

    fundmanagerdetails_excel = None
    selected_sheet = None

    if excel_candidates:
        excel_candidates = sorted(
            excel_candidates,
            key=lambda item: (-item["score"], tab3_natural_key(item["path"].name)),
        )

        best = excel_candidates[0]

        if best["score"] > 0:
            fundmanagerdetails_excel = best["path"]
            selected_sheet = best["sheet"]

    pptx = None

    if pptx_files:
        pptx = sorted(
            pptx_files,
            key=lambda p: (
                -sum(
                    10
                    for word in ["fund", "manager", "performance"]
                    if word in p.name.lower()
                ),
                tab3_natural_key(p.name),
            ),
        )[0]

    return {
        "fundmanagerdetails_excel": fundmanagerdetails_excel,
        "selected_sheet": selected_sheet,
        "pptx": pptx,
        "excel_files": excel_files,
        "pptx_files": pptx_files,
        "excel_candidates": excel_candidates,
    }


# =========================================================
# PPT HELPERS
# =========================================================

def tab3_iter_shapes(shapes):
    for shape in shapes:
        yield shape

        if hasattr(shape, "shapes"):
            yield from tab3_iter_shapes(shape.shapes)


def tab3_shape_text(shape) -> str:
    parts = []

    if getattr(shape, "has_text_frame", False):
        parts.append(shape.text)

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text)

    return "\n".join(parts)


def tab3_slide_text(slide) -> str:
    return "\n".join(
        text
        for text in (tab3_shape_text(shape) for shape in tab3_iter_shapes(slide.shapes))
        if text
    )


def tab3_all_ppt_text(prs: Presentation) -> str:
    return "\n".join(tab3_slide_text(slide) for slide in prs.slides)


def tab3_find_slides_containing(prs: Presentation, search_text: str) -> list[int]:
    needle = tab3_clean_text(search_text)

    if not needle:
        return []

    return [
        idx
        for idx, slide in enumerate(prs.slides, start=1)
        if needle in tab3_clean_text(tab3_slide_text(slide))
    ]


def tab3_extract_declared_note_counts(ppt_text: str) -> list[dict]:
    text = tab3_normalize_text(ppt_text)

    pattern = re.compile(
        r"\b("
        r"(?:Mr\.?|Ms\.?|Mrs\.?|Dr\.?)\s*"
        r"[A-Za-z][A-Za-z .'-]{1,80}?"
        r")\s+manages\s+(\d+)\s+schemes?\s+of\s+"
        r"([A-Za-z0-9& .,'()-]{2,120}?)"
        r"(?:\.|;|\n|$)",
        flags=re.IGNORECASE,
    )

    found = []
    seen = set()

    for match in pattern.finditer(text):
        manager = tab3_normalize_text(match.group(1))
        count = int(match.group(2))
        amc = tab3_normalize_text(match.group(3))
        note_text = tab3_normalize_text(match.group(0))

        if not manager:
            continue

        if len(manager.split()) > 12:
            continue

        key = (tab3_manager_match_key(manager), count, tab3_clean_text(amc))

        if key in seen:
            continue

        seen.add(key)

        found.append(
            {
                "manager": manager,
                "manager_key": tab3_manager_match_key(manager),
                "declared_count": count,
                "amc": amc,
                "note_text": note_text,
            }
        )

    return found


def tab3_get_ppt_note_for_manager(
    manager_key: str,
    declared_df: pd.DataFrame,
) -> tuple[Optional[int], str, str]:
    manager_note_rows = declared_df[
        declared_df["manager_key"] == manager_key
    ]

    if manager_note_rows.empty:
        return None, "", ""

    selected_note = manager_note_rows.iloc[0]

    return (
        int(selected_note["declared_count"]),
        selected_note["note_text"],
        selected_note["amc"],
    )


# =========================================================
# REPORT BUILDING
# =========================================================

def tab3_make_easy_report(details_df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for _, row in details_df.iterrows():
        manager_name = row.get("Manager Name", "")
        result = row.get("Result", "")
        ppt_count = row.get("PPT Says Count", "")
        excel_count = row.get("Excel Fund Manager Count", "")
        ppt_note_found = row.get("PPT Note Found", "")

        if result == "OK":
            what_to_fix = "Nothing. This manager count is correct."
        else:
            fixes = []

            if not ppt_note_found:
                fixes.append("Add the correct PPT manager count note.")
            elif str(ppt_count) != str(excel_count):
                fixes.append(f"Change PPT count from {ppt_count} to {excel_count}.")

            if not fixes:
                fixes.append("Check this manager manually.")

            what_to_fix = " ".join(fixes)

        rows.append(
            {
                "Manager Name": manager_name,
                "Result": result,
                "PPT Says Count": ppt_count,
                "Excel Fund Manager Count": excel_count,
                "What To Fix": what_to_fix,
                "Correct Note To Use": row.get("Correct Note To Use", ""),
                "PPT Note Found": ppt_note_found,
                "Schemes Counted From Excel": row.get("Schemes Counted From Excel", ""),
                "Reason": row.get("Reason", ""),
                "Manager PPT Slide": row.get("Manager PPT Slide", ""),
            }
        )

    if not rows:
        rows.append(
            {
                "Manager Name": "No data found",
                "Result": "CHECK",
                "PPT Says Count": "",
                "Excel Fund Manager Count": "",
                "What To Fix": "No manager data found. Please check Excel and PPT.",
                "Correct Note To Use": "",
                "PPT Note Found": "",
                "Schemes Counted From Excel": "",
                "Reason": "",
                "Manager PPT Slide": "",
            }
        )

    return pd.DataFrame(rows)


# =========================================================
# WORKER
# =========================================================

@dataclass
class tab3_CheckConfig:
    fundmanagerdetails_path: Path
    pptx_path: Path
    output_report_path: Path
    selected_sheet: str


class tab3_CheckerWorker(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, config: tab3_CheckConfig):
        super().__init__()
        self.config = config
        self.log_lines: list[str] = []

    def emit_log(self, message: str):
        line = message.rstrip()
        self.log_lines.append(line)
        self.log_signal.emit(line)

    def progress(self, value: int, message: str):
        self.progress_signal.emit(value, message)
        self.emit_log(f"[{value:03d}%] {message}")

    def run(self):
        try:
            cfg = self.config

            self.emit_log("")
            self.emit_log("=" * 100)
            self.emit_log("FUND MANAGER COUNT CHECK STARTED")
            self.emit_log("=" * 100)
            self.emit_log(f"Started at: {tab3_now_text()}")
            self.emit_log(f"Excel: {cfg.fundmanagerdetails_path}")
            self.emit_log(f"PPTX: {cfg.pptx_path}")

            self.progress(5, "Reading Excel...")

            df, selected_sheet, header_row, detected = tab3_read_input_excel(
                cfg.fundmanagerdetails_path,
                cfg.selected_sheet,
            )

            missing = [
                name
                for name, col in [
                    ("Scheme Name", detected["fund_name"]),
                    ("Fund Manager", detected["fund_manager"]),
                ]
                if not col
            ]

            if missing:
                raise ValueError("Missing required Excel columns: " + ", ".join(missing))

            self.emit_log(f"Selected sheet: {selected_sheet}")
            self.emit_log(f"Header row: {header_row + 1}")
            self.emit_log(f"Scheme column: {detected['fund_name']}")
            self.emit_log(f"Fund Manager column: {detected['fund_manager']}")
            self.emit_log(f"AMC column: {detected['amc'] or 'Not found, using default'}")
            self.emit_log("Excel logic: count unique schemes for each fund manager.")
            self.emit_log("PPT logic: compare only manager count statements, not scheme rows.")

            fund_col = detected["fund_name"]
            fm_col = detected["fund_manager"]
            amc_col = detected["amc"]

            self.progress(25, "Preparing fund manager rows from Excel...")

            manager_rows = []

            for idx, source_row in df.iterrows():
                fund_name = tab3_normalize_text(source_row.get(fund_col, ""))
                raw_managers = tab3_normalize_text(source_row.get(fm_col, ""))
                amc = tab3_normalize_text(source_row.get(amc_col, "")) if amc_col else tab3_DEFAULT_AMC

                if not fund_name or not raw_managers:
                    continue

                managers = tab3_split_hash_values(raw_managers)

                if not managers:
                    managers = [raw_managers]

                for manager_name in managers:
                    manager_rows.append(
                        {
                            "Manager Name": tab3_normalize_text(manager_name),
                            "Manager Key": tab3_manager_match_key(manager_name),
                            "Fund Name": fund_name,
                            "Fund Match Key": tab3_scheme_match_key(fund_name),
                            "AMC": amc or tab3_DEFAULT_AMC,
                            "Source Row": idx + header_row + 2,
                        }
                    )

            manager_df = pd.DataFrame(manager_rows)

            if manager_df.empty:
                raise ValueError("No valid manager rows found in Excel.")

            self.emit_log(f"Manager rows after # split: {len(manager_df)}")
            self.emit_log(f"Unique managers: {manager_df['Manager Key'].nunique()}")
            self.emit_log(f"Unique funds: {manager_df['Fund Match Key'].nunique()}")

            self.progress(40, "Calculating unique scheme count per manager...")

            summary_rows = []

            for manager_key, group in manager_df.groupby("Manager Key", dropna=False):
                unique_fund_rows = group.drop_duplicates("Fund Match Key")
                fund_count = int(unique_fund_rows["Fund Match Key"].nunique())

                display_names = [
                    tab3_normalize_text(value)
                    for value in group["Manager Name"].dropna().unique().tolist()
                    if tab3_normalize_text(value)
                ]
                manager_name = display_names[0] if display_names else manager_key

                amc_values = [
                    tab3_normalize_text(value)
                    for value in unique_fund_rows["AMC"].dropna().unique().tolist()
                    if tab3_normalize_text(value)
                ]
                amc = amc_values[0] if amc_values else tab3_DEFAULT_AMC

                counted_funds = unique_fund_rows["Fund Name"].dropna().tolist()

                summary_rows.append(
                    {
                        "Manager Name": manager_name,
                        "Manager Key": manager_key,
                        "AMC": amc,
                        "Excel Fund Manager Count": fund_count,
                        "Schemes Counted From Excel": "; ".join(dict.fromkeys(counted_funds)),
                    }
                )

            summary_df = pd.DataFrame(summary_rows).sort_values(
                "Manager Name",
                key=lambda s: s.map(tab3_natural_key),
            )

            self.progress(55, "Reading PPT manager count statements...")

            prs = Presentation(str(cfg.pptx_path))
            ppt_text = tab3_all_ppt_text(prs)
            ppt_clean = tab3_clean_text(ppt_text)

            declared_df = pd.DataFrame(tab3_extract_declared_note_counts(ppt_text))

            if declared_df.empty:
                declared_df = pd.DataFrame(
                    columns=["manager", "manager_key", "declared_count", "amc", "note_text"]
                )

            self.emit_log(f"Declared manager count statements found in PPT: {len(declared_df)}")

            self.progress(70, "Comparing Excel counts with PPT statements...")

            detail_rows = []

            for _, row in summary_df.iterrows():
                manager_name = row["Manager Name"]
                manager_key = row["Manager Key"]
                amc = row["AMC"] or tab3_DEFAULT_AMC
                excel_count = int(row["Excel Fund Manager Count"])

                correct_note = tab3_expected_note_text(manager_name, excel_count, amc)

                ppt_count, ppt_note_found, ppt_amc = tab3_get_ppt_note_for_manager(
                    manager_key=manager_key,
                    declared_df=declared_df,
                )

                manager_slides = tab3_find_slides_containing(prs, manager_name)

                note_found = bool(ppt_note_found)
                count_ok = ppt_count == excel_count
                amc_ok = tab3_amc_matches(amc, ppt_amc)

                overall_ok = note_found and count_ok and amc_ok
                result = "OK" if overall_ok else "FIX"

                reason_parts = []

                if manager_key not in ppt_clean:
                    reason_parts.append("Manager name not found in PPT")

                if not note_found:
                    reason_parts.append("PPT manager count statement not found")

                if note_found and not count_ok:
                    reason_parts.append("PPT count and Excel count are different")

                if note_found and not amc_ok:
                    reason_parts.append("AMC in PPT statement is different from Excel/default AMC")

                if not reason_parts:
                    reason_parts.append("Everything matched")

                detail_rows.append(
                    {
                        "Manager Name": manager_name,
                        "Result": result,
                        "PPT Says Count": ppt_count if ppt_count is not None else "Not found",
                        "Excel Fund Manager Count": excel_count,
                        "Correct Note To Use": correct_note,
                        "PPT Note Found": ppt_note_found,
                        "Schemes Counted From Excel": row["Schemes Counted From Excel"],
                        "Reason": "; ".join(reason_parts),
                        "Manager PPT Slide": ", ".join(map(str, manager_slides)) if manager_slides else "Not found",
                    }
                )

                self.emit_log(
                    f"[{result}] {manager_name}: PPT={ppt_count}, Excel={excel_count}"
                )

            details_df = pd.DataFrame(detail_rows)
            easy_report_df = tab3_make_easy_report(details_df)

            ok_count = int((easy_report_df["Result"] == "OK").sum())
            fix_count = int((easy_report_df["Result"] == "FIX").sum())

            self.progress(90, "Writing Easy Report...")

            tab3_write_excel_report(cfg.output_report_path, easy_report_df)

            self.progress(100, "Completed")

            self.finished_signal.emit(
                True,
                f"Fund manager count check completed.\n\n"
                f"OK: {ok_count}\n"
                f"Need Fix: {fix_count}\n\n"
                f"Easy Report saved at:\n{cfg.output_report_path}",
            )

        except Exception as exc:
            self.emit_log("")
            self.emit_log("=" * 100)
            self.emit_log("ERROR")
            self.emit_log("=" * 100)
            self.emit_log(str(exc))
            self.finished_signal.emit(False, str(exc))


# =========================================================
# GUI
# =========================================================

class tab3_FundManagerPPTChecker(QWidget):
    def __init__(self, state: Any | None = None):
        super().__init__()
        self.state = state

        self.setWindowTitle(tab3_APP_TITLE)
        self.resize(1350, 780)

        self.source_folder: Optional[Path] = None
        self.fundmanagerdetails_path: Optional[Path] = None
        self.pptx_path: Optional[Path] = None
        self.output_report_path: Optional[Path] = None
        self.selected_sheet: str = ""
        self.worker: Optional[tab3_CheckerWorker] = None

        self.build_ui()
        self._apply_default_output_report_path()


    def _app_run_folder(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parents[1]

    def _resolve_default_output_folder(self) -> Path:
        state_folder = getattr(self.state, "tab3_output_folder", None) if getattr(self, "state", None) is not None else None
        if state_folder:
            folder = Path(state_folder)
        else:
            output_root = getattr(self.state, "output_root", None) if getattr(self, "state", None) is not None else None
            folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "03_fund_manager_count"
        folder.mkdir(parents=True, exist_ok=True)
        if getattr(self, "state", None) is not None:
            setattr(self.state, "tab3_output_folder", folder)
        return folder

    def _apply_default_output_report_path(self) -> None:
        folder = self._resolve_default_output_folder()
        if not self.output_report_path:
            self.output_report_path = folder / "easy_fund_manager_count_report.xlsx"
            self.output_report_edit.setText(str(self.output_report_path))

    def build_ui(self):
        root = QHBoxLayout(self)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(splitter)

        log_panel = QWidget()
        log_layout = QVBoxLayout(log_panel)

        log_title = QLabel("Process Logs")
        log_title.setStyleSheet("font-weight: bold; font-size: 15px;")
        log_layout.addWidget(log_title)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        log_layout.addWidget(self.log_box)

        clear_log = QPushButton("Clear Log")
        clear_log.clicked.connect(self.log_box.clear)
        log_layout.addWidget(clear_log)

        splitter.addWidget(log_panel)

        main_panel = QWidget()
        main_layout = QVBoxLayout(main_panel)

        title = QLabel("Fund Manager Count PPT Checker")
        title.setStyleSheet("font-weight: bold; font-size: 18px;")
        main_layout.addWidget(title)

        desc = QLabel(
            "This tool checks only manager count statements in PPT, like "
            "'Mr. Gautam Kaul manages 14 schemes of Bandhan Mutual Fund.' "
            "It does not compare scheme/fund rows in PPT."
        )
        desc.setWordWrap(True)
        main_layout.addWidget(desc)

        main_layout.addWidget(self.create_source_group())
        main_layout.addWidget(self.create_file_group())
        main_layout.addWidget(self.create_output_group())
        main_layout.addWidget(self.create_progress_group())
        main_layout.addStretch()

        splitter.addWidget(main_panel)
        splitter.setSizes([420, 930])

    def create_source_group(self):
        source_group = QGroupBox("1. Source Folder / Auto Detection")
        source_grid = QGridLayout(source_group)

        self.source_folder_edit = QLineEdit()
        self.source_folder_edit.setReadOnly(True)

        btn_source = QPushButton("Browse Source Folder...")
        btn_source.clicked.connect(self.browse_source_folder)

        btn_detect = QPushButton("Auto Detect Files")
        btn_detect.clicked.connect(self.auto_detect_files)

        source_grid.addWidget(QLabel("Source Folder:"), 0, 0)
        source_grid.addWidget(self.source_folder_edit, 0, 1)
        source_grid.addWidget(btn_source, 0, 2)
        source_grid.addWidget(btn_detect, 0, 3)

        return source_group

    def create_file_group(self):
        file_group = QGroupBox("2. Files")
        grid = QGridLayout(file_group)

        self.fundmanagerdetails_edit = QLineEdit()
        self.fundmanagerdetails_edit.setReadOnly(True)

        self.pptx_edit = QLineEdit()
        self.pptx_edit.setReadOnly(True)

        btn_excel = QPushButton("Browse fundmanagerdetails Excel...")
        btn_excel.clicked.connect(self.browse_fundmanagerdetails)

        btn_pptx = QPushButton("Browse PPT...")
        btn_pptx.clicked.connect(self.browse_pptx)

        grid.addWidget(QLabel("fundmanagerdetails Excel:"), 0, 0)
        grid.addWidget(self.fundmanagerdetails_edit, 0, 1)
        grid.addWidget(btn_excel, 0, 2)

        grid.addWidget(QLabel("PPTX:"), 1, 0)
        grid.addWidget(self.pptx_edit, 1, 1)
        grid.addWidget(btn_pptx, 1, 2)

        return file_group

    def create_output_group(self):
        group = QGroupBox("3. Output / Run")
        grid = QGridLayout(group)

        self.output_report_edit = QLineEdit()
        self.output_report_edit.setReadOnly(True)

        btn_output = QPushButton("Save Easy Report As...")
        btn_output.clicked.connect(self.browse_report_output)

        self.run_button = QPushButton("Run Fund Manager Count Check")
        self.run_button.clicked.connect(self.run_check)

        grid.addWidget(QLabel("Easy Report:"), 0, 0)
        grid.addWidget(self.output_report_edit, 0, 1)
        grid.addWidget(btn_output, 0, 2)
        grid.addWidget(self.run_button, 1, 2)

        return group

    def create_progress_group(self):
        progress_group = QGroupBox("Progress")
        p_layout = QVBoxLayout(progress_group)

        self.status_label = QLabel("Idle")

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)

        p_layout.addWidget(self.status_label)
        p_layout.addWidget(self.progress_bar)

        return progress_group

    def log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def set_default_output(self, base_folder: Path):
        if not self.output_report_path:
            folder = self._resolve_default_output_folder()
            self.output_report_path = folder / "easy_fund_manager_count_report.xlsx"
            self.output_report_edit.setText(str(self.output_report_path))

    def browse_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Source Folder")

        if folder:
            self.source_folder = Path(folder)
            self.source_folder_edit.setText(str(self.source_folder))
            self.auto_detect_files()

    def auto_detect_files(self):
        if not self.source_folder:
            folder_text = self.source_folder_edit.text().strip()

            if folder_text:
                self.source_folder = Path(folder_text)

        if not self.source_folder or not self.source_folder.exists():
            QMessageBox.warning(
                self,
                "Missing Source Folder",
                "Please select a valid source folder first.",
            )
            return

        self.log("")
        self.log("=" * 100)
        self.log("SOURCE FOLDER FILE DETECTION")
        self.log("=" * 100)
        self.log(f"Source folder: {self.source_folder}")

        detected = tab3_auto_detect_files_from_folder(self.source_folder)

        self.log(f"Excel files found: {len(detected['excel_files'])}")

        for candidate in detected.get("excel_candidates", []):
            path = candidate["path"]
            score = candidate["score"]
            sheet = candidate["sheet"]
            self.log(f"  - Excel: {path.name} | score: {score} | best sheet: {sheet}")

        self.log(f"PPTX files found: {len(detected['pptx_files'])}")

        for p in detected["pptx_files"]:
            self.log(f"  - PPTX: {p.name}")

        if detected["fundmanagerdetails_excel"]:
            self.fundmanagerdetails_path = detected["fundmanagerdetails_excel"]
            self.selected_sheet = detected.get("selected_sheet") or ""
            self.fundmanagerdetails_edit.setText(str(self.fundmanagerdetails_path))
            self.log(f"Selected fundmanagerdetails Excel: {self.fundmanagerdetails_path.name}")
            self.log(f"Selected sheet: {self.selected_sheet}")
        else:
            self.log("Correct fundmanagerdetails Excel not detected. Browse manually.")

        if detected["pptx"]:
            self.pptx_path = detected["pptx"]
            self.pptx_edit.setText(str(self.pptx_path))
            self.log(f"Selected PPTX: {self.pptx_path.name}")
        else:
            self.log("PPTX not detected. Browse manually.")

        self.set_default_output(self.source_folder)
        self.log(f"Easy Report output: {self.output_report_path}")

    def browse_fundmanagerdetails(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select fundmanagerdetails Excel",
            str(self.source_folder or Path.cwd()),
            "Excel Files (*.xlsx *.xlsm *.xls)",
        )

        if file_path:
            self.fundmanagerdetails_path = Path(file_path)
            score, sheet, _, _ = tab3_fundmanagerdetails_content_score(self.fundmanagerdetails_path)
            self.selected_sheet = sheet or ""
            self.fundmanagerdetails_edit.setText(str(self.fundmanagerdetails_path))
            self.set_default_output(self.fundmanagerdetails_path.parent)
            self.log(f"Manual fundmanagerdetails selected: {self.fundmanagerdetails_path}")
            self.log(f"Detected best sheet: {self.selected_sheet}")
            self.log(f"Excel score: {score}")

    def browse_pptx(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select FUND MANAGER PERFORMANCE PPT",
            str(self.source_folder or Path.cwd()),
            "PowerPoint Files (*.pptx)",
        )

        if file_path:
            self.pptx_path = Path(file_path)
            self.pptx_edit.setText(str(self.pptx_path))
            self.set_default_output(self.pptx_path.parent)
            self.log(f"Manual PPT selected: {self.pptx_path}")

    def browse_report_output(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Easy Report As",
            str(self.output_report_path or (self._resolve_default_output_folder() / "easy_fund_manager_count_report.xlsx")),
            "Excel Files (*.xlsx)",
        )

        if file_path:
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"

            self.output_report_path = Path(file_path)
            self.output_report_edit.setText(str(self.output_report_path))
            self.log(f"Manual report output selected: {self.output_report_path}")

    def validate_inputs(self) -> tuple[bool, str]:
        if not self.fundmanagerdetails_path or not self.fundmanagerdetails_path.exists():
            return False, "Please select a valid fundmanagerdetails Excel file."

        if not self.pptx_path or not self.pptx_path.exists():
            return False, "Please select a valid PPTX file."

        if not self.output_report_path:
            return False, "Please select output Excel report path."

        if self.output_report_path.suffix.lower() != ".xlsx":
            self.output_report_path = self.output_report_path.with_suffix(".xlsx")
            self.output_report_edit.setText(str(self.output_report_path))

        return True, ""

    def run_check(self):
        ok, msg = self.validate_inputs()

        if not ok:
            QMessageBox.warning(self, "Validation Error", msg)
            return

        cfg = tab3_CheckConfig(
            fundmanagerdetails_path=self.fundmanagerdetails_path,
            pptx_path=self.pptx_path,
            output_report_path=self.output_report_path,
            selected_sheet=self.selected_sheet,
        )

        self.progress_bar.setValue(0)
        self.status_label.setText("Starting fund manager count check...")
        self.run_button.setEnabled(False)

        self.worker = tab3_CheckerWorker(cfg)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def update_progress(self, value: int, message: str):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)

    def on_finished(self, success: bool, message: str):
        self.run_button.setEnabled(True)

        if success:
            self.progress_bar.setValue(100)
            self.status_label.setText("Completed")
            QMessageBox.information(self, "Completed", message)
        else:
            self.status_label.setText("Failed")
            QMessageBox.critical(self, "Failed", message)



tab3_Tab3Widget = tab3_FundManagerPPTChecker

def tab3_main():
    app = QApplication(sys.argv)
    window = tab3_FundManagerPPTChecker()
    window.show()
    sys.exit(app.exec())


if False and __name__ == "__main__":
    tab3_main()


# ============================================================
# COMBINED TAB UI - Tab 5 style section tabs
# ============================================================

class Section1PerformanceTableNoteWindow(tab1_PerformanceTableNoteWindow):
    """Section 1 wrapper.

    Keeps the original Section 1 logic unchanged, and only notifies the combined
    parent when the output PPTX is created so Section 2 can use that PPTX
    automatically.
    """

    def __init__(self, state: Any | None = None, output_ready_callback=None):
        self.output_ready_callback = output_ready_callback
        super().__init__(state)

    def on_finished(self, result):
        super().on_finished(result)
        output_path = result.get("output_path") if isinstance(result, dict) else None
        if output_path:
            output_path = Path(output_path)
            if self.state is not None:
                setattr(self.state, "tab1_latest_output_pptx", output_path)
                setattr(self.state, "section1_output_pptx", output_path)
            if callable(self.output_ready_callback):
                self.output_ready_callback(output_path)


class Section2FundManagerPPTChecker(tab3_FundManagerPPTChecker):
    """Section 2 wrapper.

    Updates only the GUI/input flow:
    - PPT input can be filled automatically from Section 1 output.
    - Manual PPT input is still available through a browse button.
    - Fundmanagerdetails Excel remains manually selectable / auto-detectable as before.
    """

    def __init__(self, state: Any | None = None):
        super().__init__(state)
        self.try_use_section1_output_ppt(silent=True)

    def create_file_group(self):
        file_group = QGroupBox("2. Files")
        grid = QGridLayout(file_group)

        self.fundmanagerdetails_edit = QLineEdit()
        self.fundmanagerdetails_edit.setReadOnly(True)

        self.pptx_edit = QLineEdit()
        self.pptx_edit.setReadOnly(True)
        self.pptx_edit.setPlaceholderText("Auto-filled from Section 1 output, or browse manually")

        btn_excel = QPushButton("Browse fundmanagerdetails Excel...")
        btn_excel.clicked.connect(self.browse_fundmanagerdetails)

        btn_use_section1 = QPushButton("Use Section 1 Output PPT")
        btn_use_section1.clicked.connect(lambda: self.try_use_section1_output_ppt(silent=False))

        btn_pptx = QPushButton("Browse Manual PPT...")
        btn_pptx.clicked.connect(self.browse_pptx)

        grid.addWidget(QLabel("fundmanagerdetails Excel:"), 0, 0)
        grid.addWidget(self.fundmanagerdetails_edit, 0, 1)
        grid.addWidget(btn_excel, 0, 2, 1, 2)

        grid.addWidget(QLabel("PPTX Input:"), 1, 0)
        grid.addWidget(self.pptx_edit, 1, 1)
        grid.addWidget(btn_use_section1, 1, 2)
        grid.addWidget(btn_pptx, 1, 3)

        help_text = QLabel("PPTX defaults to Section 1 output after Section 1 completes. Use Browse Manual PPT only if you want to override it.")
        help_text.setWordWrap(True)
        grid.addWidget(help_text, 2, 1, 1, 3)

        return file_group

    def get_section1_output_ppt(self) -> Path | None:
        candidates = []
        if self.state is not None:
            for attr in ["tab1_latest_output_pptx", "section1_output_pptx"]:
                value = getattr(self.state, attr, None)
                if value:
                    candidates.append(Path(value))

            # Fallback: find newest PPTX inside Section 1 output folder.
            folder = getattr(self.state, "tab1_section1_performance_table_note_folder", None)
            if not folder:
                folder = getattr(self.state, "tab1_output_folder", None)
            if folder and Path(folder).exists():
                pptx_files = [p for p in Path(folder).glob("*.pptx") if not p.name.startswith("~$")]
                if pptx_files:
                    candidates.append(max(pptx_files, key=lambda p: p.stat().st_mtime))

        for candidate in candidates:
            if candidate.exists() and candidate.suffix.lower() == ".pptx":
                return candidate
        return None

    def set_section1_output_ppt(self, pptx_path):
        pptx_path = Path(pptx_path)
        if pptx_path.exists():
            self.pptx_path = pptx_path
            self.pptx_edit.setText(str(pptx_path))
            self.set_default_output(pptx_path.parent)
            self.log(f"[Section2] PPT input set from Section 1 output: {pptx_path}")

    def try_use_section1_output_ppt(self, silent: bool = False):
        pptx_path = self.get_section1_output_ppt()
        if pptx_path:
            self.set_section1_output_ppt(pptx_path)
            return True
        if not silent:
            QMessageBox.information(
                self,
                "Section 1 Output Not Found",
                "Section 1 output PPT was not found yet. Run Section 1 first, or use Browse Manual PPT.",
            )
            self.log("[Section2] Section 1 output PPT not found. Use manual PPT browse if needed.")
        return False

    def browse_pptx(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Manual PPT Input for Fund Manager Count",
            str(self.source_folder or Path.cwd()),
            "PowerPoint Files (*.pptx)",
        )

        if file_path:
            self.pptx_path = Path(file_path)
            self.pptx_edit.setText(str(self.pptx_path))
            self.set_default_output(self.pptx_path.parent)
            self.log(f"[Section2] Manual PPT selected: {self.pptx_path}")

    def run_check(self):
        # If user did not manually select PPT, try to auto-load Section 1 output.
        if not self.pptx_path or not self.pptx_path.exists():
            self.try_use_section1_output_ppt(silent=True)
        super().run_check()


class CombinedTab1Tab3Widget(QWidget):
    """Combined widget with two independently-run sections.

    Section 1 creates the performance-table notes-updated PPTX.
    Section 2 uses Section 1 output PPTX automatically as PPT input, while still
    allowing manual PPT selection.
    """

    def __init__(self, state=None):
        super().__init__()
        self.state = state

        layout = QVBoxLayout(self)

        title = QLabel("Performance Table Note / Fund Manager Count")
        title.setStyleSheet("font-weight: bold; font-size: 18px;")
        layout.addWidget(title)

        self.section_tabs = QTabWidget()
        self.section1_widget = Section1PerformanceTableNoteWindow(state, self.on_section1_output_ready)
        self.section2_widget = Section2FundManagerPPTChecker(state)

        self.section_tabs.addTab(self.section1_widget, "Section 1 - Performance Table Note")
        self.section_tabs.addTab(self.section2_widget, "Section 2 - Fund Manager Count")
        layout.addWidget(self.section_tabs, stretch=1)

    def on_section1_output_ready(self, output_path: Path):
        if self.state is not None:
            setattr(self.state, "tab1_latest_output_pptx", Path(output_path))
            setattr(self.state, "section1_output_pptx", Path(output_path))
        if hasattr(self, "section2_widget"):
            self.section2_widget.set_section1_output_ppt(output_path)

# Friendly aliases for loaders that expect tab widget class names.
Tab1Widget = CombinedTab1Tab3Widget
Tab3Widget = CombinedTab1Tab3Widget
PerformanceTableNoteWidget = CombinedTab1Tab3Widget
FundManagerCountWidget = CombinedTab1Tab3Widget
CombinedPerformanceFundManagerWidget = CombinedTab1Tab3Widget


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Combined Tab 1 + Tab 3 Automation")
        self.resize(1435, 850)
        self.setCentralWidget(CombinedTab1Tab3Widget())


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
