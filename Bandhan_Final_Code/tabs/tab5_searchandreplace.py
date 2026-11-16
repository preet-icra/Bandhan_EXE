from __future__ import annotations

import re
import sys
import traceback
from pathlib import Path
from typing import Any, Callable

from PyQt6.QtCore import QObject, QRunnable, QThreadPool, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QTextEdit,
    QMessageBox,
    QGroupBox,
    QCheckBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QAbstractItemView,
    QScrollArea,
    QSplitter,
    QProgressBar,
)

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


# =========================================================
# WORKER
# =========================================================
class WorkerSignals(QObject):
    started = pyqtSignal()
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)


class FunctionWorker(QRunnable):
    """Run a normal Python function inside Qt thread pool."""

    def __init__(self, func: Callable[..., Any], *args: Any, **kwargs: Any) -> None:
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

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
# SEARCH / REPLACE CORE LOGIC
# =========================================================
def _norm(value: str, case_sensitive: bool) -> str:
    return value if case_sensitive else value.lower()


def _has_search(text: str, search_keyword: str, case_sensitive: bool) -> bool:
    if not text or not search_keyword:
        return False

    return _norm(search_keyword, case_sensitive) in _norm(text, case_sensitive)


def _replace_text(
    source_text: str,
    search_keyword: str,
    replace_with: str,
    case_sensitive: bool,
) -> tuple[str, int]:
    """
    Returns:
        tuple[str, int]: updated text and replacement count.
    """
    if not source_text or not search_keyword:
        return source_text, 0

    if case_sensitive:
        return (
            source_text.replace(search_keyword, replace_with),
            source_text.count(search_keyword),
        )

    source_lower = source_text.lower()
    search_lower = search_keyword.lower()

    result: list[str] = []
    start = 0
    count = 0

    while True:
        index = source_lower.find(search_lower, start)

        if index == -1:
            result.append(source_text[start:])
            break

        result.append(source_text[start:index])
        result.append(replace_with)

        start = index + len(search_keyword)
        count += 1

    return "".join(result), count


def _shape_type_name(shape) -> str:
    if getattr(shape, "has_table", False):
        return "table"

    if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
        return "textbox"

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "group"

    return "shape"


def _is_blacklisted(shape, blacklist: list[str], case_sensitive: bool) -> bool:
    """
    Blacklist supports:
    - Exact shape name
    - Object type:
        table
        textbox
        group
        shape
    """
    shape_name = getattr(shape, "name", "")
    obj_type = _shape_type_name(shape)

    normalized_blacklist = [
        _norm(item.strip(), case_sensitive)
        for item in blacklist
        if item.strip()
    ]

    return (
        _norm(shape_name, case_sensitive) in normalized_blacklist
        or _norm(obj_type, case_sensitive) in normalized_blacklist
    )


def _apply_pairs_to_text(
    source_text: str,
    pairs: list[tuple[str, str]],
    case_sensitive: bool,
) -> tuple[str, int]:
    """
    Apply search/replace pairs in order.

    Returns:
        tuple[str, int]: updated text and total replacement count.
    """
    updated = source_text
    total_count = 0

    for search_keyword, replace_with in pairs:
        if not search_keyword:
            continue

        if _has_search(updated, search_keyword, case_sensitive):
            updated, count = _replace_text(
                updated,
                search_keyword,
                replace_with,
                case_sensitive,
            )
            total_count += count

    return updated, total_count


def _replace_in_paragraph(
    paragraph,
    pairs: list[tuple[str, str]],
    case_sensitive: bool,
) -> int:
    """
    Tries to preserve formatting.

    Best case:
        Replacement happens inside individual runs.

    Fallback:
        If search text is split across runs, full paragraph text is replaced
        into the first run and other runs are cleared.
    """
    replacements = 0

    # First try per-run replacement to preserve formatting.
    for run in paragraph.runs:
        new_text, count = _apply_pairs_to_text(
            run.text,
            pairs,
            case_sensitive,
        )

        if count:
            run.text = new_text
            replacements += count

    if replacements:
        return replacements

    # Fallback for text split across runs.
    full_text = paragraph.text
    new_text, replacements = _apply_pairs_to_text(
        full_text,
        pairs,
        case_sensitive,
    )

    if not replacements:
        return 0

    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = new_text

        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run().text = new_text

    return replacements


def _replace_in_text_frame(
    text_frame,
    pairs: list[tuple[str, str]],
    case_sensitive: bool,
) -> int:
    replacements = 0

    for paragraph in text_frame.paragraphs:
        replacements += _replace_in_paragraph(
            paragraph,
            pairs,
            case_sensitive,
        )

    return replacements


def _replace_in_table(
    table,
    pairs: list[tuple[str, str]],
    case_sensitive: bool,
) -> int:
    replacements = 0

    for row in table.rows:
        for cell in row.cells:
            replacements += _replace_in_text_frame(
                cell.text_frame,
                pairs,
                case_sensitive,
            )

    return replacements


def _process_shape_for_search_replace(
    shape,
    pairs: list[tuple[str, str]],
    blacklist: list[str],
    case_sensitive: bool,
) -> int:
    if _is_blacklisted(shape, blacklist, case_sensitive):
        return 0

    replacements = 0

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child_shape in shape.shapes:
            replacements += _process_shape_for_search_replace(
                child_shape,
                pairs,
                blacklist,
                case_sensitive,
            )

        return replacements

    if getattr(shape, "has_table", False):
        replacements += _replace_in_table(
            shape.table,
            pairs,
            case_sensitive,
        )
        return replacements

    if getattr(shape, "has_text_frame", False):
        replacements += _replace_in_text_frame(
            shape.text_frame,
            pairs,
            case_sensitive,
        )

    return replacements


def run_search_replace_action(
    input_pptx: Path,
    output_pptx: Path,
    pairs: list[tuple[str, str]],
    blacklist: list[str] | None = None,
    case_sensitive: bool = False,
    progress_cb=None,
) -> Path:
    """
    Search/replace text inside a single PPTX.

    Handles:
    - normal textboxes
    - shapes with text frames
    - tables
    - grouped shapes

    Preserves formatting as much as possible.
    """

    def log(msg: str) -> None:
        if progress_cb:
            progress_cb(msg)

    input_pptx = Path(input_pptx)
    output_pptx = Path(output_pptx)

    if not input_pptx.exists():
        raise FileNotFoundError(f"Input PPTX not found: {input_pptx}")

    if not pairs:
        raise ValueError("At least one search/replace pair is required.")

    blacklist = blacklist or []

    prs = Presentation(str(input_pptx))
    total_replacements = 0
    total_slides = max(1, len(prs.slides))

    log(f"[SearchReplace] Loading: {input_pptx}")
    log(f"[SearchReplace] Case sensitive: {case_sensitive}")
    log(f"[SearchReplace] Blacklist: {blacklist}")
    log(f"[SearchReplace] Pairs: {pairs}")
    log("[SearchReplace] Progress: 0%")

    for slide_no, slide in enumerate(prs.slides, start=1):
        slide_replacements = 0

        for shape in slide.shapes:
            slide_replacements += _process_shape_for_search_replace(
                shape,
                pairs,
                blacklist,
                case_sensitive,
            )

        total_replacements += slide_replacements

        pct = int((slide_no / total_slides) * 100)

        log(
            f"[SearchReplace] Slide {slide_no}: "
            f"{slide_replacements} replacement(s)"
        )
        log(f"[SearchReplace] Progress: {pct}%")

    output_pptx.parent.mkdir(parents=True, exist_ok=True)

    # Prevent silent failure if the output file is already open in PowerPoint.
    if output_pptx.exists():
        try:
            output_pptx.unlink()
        except PermissionError:
            raise PermissionError(
                f"Cannot overwrite output file because it may be open:\n"
                f"{output_pptx}\n\n"
                "Please close the PPTX file and try again."
            )

    prs.save(str(output_pptx))

    log(f"[SearchReplace] Total replacements: {total_replacements}")
    log(f"[SearchReplace] Saved: {output_pptx}")
    log("[SearchReplace] Progress: 100%")

    return output_pptx


def run_search_replace_batch_action(
    jobs: list[tuple[Path, Path]],
    pairs: list[tuple[str, str]],
    blacklist: list[str] | None = None,
    case_sensitive: bool = False,
    progress_cb=None,
) -> list[Path]:
    """
    Batch Search/Replace action.

    Args:
        jobs:
            List of:
                (input_pptx, output_pptx)

        pairs:
            List of:
                (search_text, replace_with)

        blacklist:
            Optional list of shape names or object types to skip.
            Supported values:
                table, textbox, shape, group

        case_sensitive:
            Whether search should be case-sensitive.
    """

    def log(msg: str) -> None:
        if progress_cb:
            progress_cb(msg)

    if not jobs:
        raise ValueError("At least one PPTX file is required.")

    completed_outputs: list[Path] = []
    total_jobs = len(jobs)

    log(f"[SearchReplace-Batch] Total PPTX files: {total_jobs}")
    log("[SearchReplace-Batch] Progress: 0%")

    for index, (input_pptx, output_pptx) in enumerate(jobs, start=1):
        input_pptx = Path(input_pptx)
        output_pptx = Path(output_pptx)

        log("=" * 70)
        log(f"[SearchReplace-Batch] File {index}/{total_jobs}")
        log(f"[SearchReplace-Batch] Input:  {input_pptx}")
        log(f"[SearchReplace-Batch] Output: {output_pptx}")

        run_search_replace_action(
            input_pptx=input_pptx,
            output_pptx=output_pptx,
            pairs=pairs,
            blacklist=blacklist,
            case_sensitive=case_sensitive,
            progress_cb=progress_cb,
        )

        completed_outputs.append(output_pptx)

        pct = int((index / total_jobs) * 100)

        log(f"[SearchReplace-Batch] Completed {index}/{total_jobs}")
        log(f"[SearchReplace-Batch] Progress: {pct}%")

    log("[SearchReplace-Batch] All files completed.")
    log("[SearchReplace-Batch] Progress: 100%")

    return completed_outputs


# =========================================================
# UI WIDGET
# =========================================================
class SearchReplaceWidget(QWidget):
    """
    Standalone Search / Replace tab.

    Features:
    - Add one or more PPTX files.
    - Add multiple search/replace pairs.
    - Optional blacklist.
    - Case-sensitive option.
    - Batch processing.
    """

    def __init__(
        self,
        output_folder: str | Path | None = None,
        state: Any | None = None,
    ):
        super().__init__()

        self.state = state

        self.thread_pool = QThreadPool.globalInstance()
        self.worker: FunctionWorker | None = None

        if output_folder:
            self.default_output_folder = Path(output_folder)
        else:
            self.default_output_folder = self._resolve_default_output_folder()

        self._build_ui()
        self._prefill_defaults()

    # =====================================================
    # UI
    # =====================================================

    def _app_run_folder(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent

        return Path(__file__).resolve().parent

    def _resolve_default_output_folder(self) -> Path:
        state = getattr(self, "state", None)

        state_folder = getattr(state, "tab6_output_folder", None) if state is not None else None

        if state_folder:
            folder = Path(state_folder)
        else:
            output_root = getattr(state, "output_root", None) if state is not None else None
            folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "06_search_replace"

        folder.mkdir(parents=True, exist_ok=True)

        if state is not None:
            setattr(state, "tab6_output_folder", folder)

        return folder

    def _build_ui(self) -> None:
        root = QHBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(10)

        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter)

        # LEFT LOG PANEL
        left_widget = QWidget()
        left_widget.setMinimumWidth(280)
        left_widget.setMaximumWidth(380)

        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(8)
        left_layout.setContentsMargins(0, 0, 0, 0)

        title = QLabel("Search / Replace PPTX")
        title.setStyleSheet("font-weight: bold; font-size: 14px;")
        left_layout.addWidget(title)

        info = QLabel(
            "Search and replace text inside PPTX files.\n"
            "Supports textboxes, tables, normal shapes, and grouped shapes."
        )
        info.setWordWrap(True)
        left_layout.addWidget(info)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Logs will appear here...")
        left_layout.addWidget(self.log_box, stretch=1)

        self.btn_clear_log = QPushButton("Clear Log")
        left_layout.addWidget(self.btn_clear_log)

        # RIGHT PANEL
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)

        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setSpacing(10)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # OUTPUT FOLDER
        output_group = QGroupBox("Output Folder")
        output_grid = QGridLayout(output_group)

        self.output_folder_edit = QLineEdit()
        self.btn_browse_output = QPushButton("Browse...")

        output_grid.addWidget(QLabel("Output Folder:"), 0, 0)
        output_grid.addWidget(self.output_folder_edit, 0, 1)
        output_grid.addWidget(self.btn_browse_output, 0, 2)

        right_layout.addWidget(output_group)

        # PPT FILES
        ppt_group = QGroupBox("PPTX Files")
        ppt_layout = QVBoxLayout(ppt_group)

        ppt_help = QLabel(
            "Tick PPTX files to process. If none are ticked, all rows are processed."
        )
        ppt_help.setWordWrap(True)
        ppt_layout.addWidget(ppt_help)

        ppt_btn_row = QHBoxLayout()

        self.btn_add_ppts = QPushButton("Add PPTX Files...")
        self.btn_remove_ppts = QPushButton("Remove Selected")
        self.btn_clear_ppts = QPushButton("Clear PPTX List")

        ppt_btn_row.addWidget(self.btn_add_ppts)
        ppt_btn_row.addWidget(self.btn_remove_ppts)
        ppt_btn_row.addWidget(self.btn_clear_ppts)
        ppt_btn_row.addStretch(1)

        ppt_layout.addLayout(ppt_btn_row)

        self.ppt_table = QTableWidget(0, 3)
        self.ppt_table.setHorizontalHeaderLabels(["Run?", "Input PPTX", "Output PPTX Name"])
        self.ppt_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.ppt_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.ppt_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.ppt_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.ppt_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        ppt_layout.addWidget(self.ppt_table)

        right_layout.addWidget(ppt_group)

        # PAIRS
        pair_group = QGroupBox("Search / Replace Pairs")
        pair_layout = QVBoxLayout(pair_group)

        self.pair_table = QTableWidget(0, 2)
        self.pair_table.setHorizontalHeaderLabels(["Search Text", "Replace With"])
        self.pair_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.pair_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.pair_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.pair_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        pair_layout.addWidget(self.pair_table)

        pair_btn_row = QHBoxLayout()

        self.btn_add_pair = QPushButton("Add Pair")
        self.btn_remove_pair = QPushButton("Remove Selected Pair")
        self.btn_clear_pairs = QPushButton("Clear Pairs")

        pair_btn_row.addWidget(self.btn_add_pair)
        pair_btn_row.addWidget(self.btn_remove_pair)
        pair_btn_row.addWidget(self.btn_clear_pairs)
        pair_btn_row.addStretch(1)

        pair_layout.addLayout(pair_btn_row)

        right_layout.addWidget(pair_group)

        # OPTIONS
        options_group = QGroupBox("Options")
        options_grid = QGridLayout(options_group)

        self.case_sensitive_chk = QCheckBox("Case Sensitive")
        self.case_sensitive_chk.setChecked(False)

        self.blacklist_edit = QLineEdit("")
        self.blacklist_edit.setPlaceholderText(
            "Comma-separated shape names or types: table, textbox, shape, group"
        )

        options_grid.addWidget(self.case_sensitive_chk, 0, 0, 1, 2)
        options_grid.addWidget(QLabel("Blacklist:"), 1, 0)
        options_grid.addWidget(self.blacklist_edit, 1, 1)

        right_layout.addWidget(options_group)

        # PROGRESS
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)

        self.progress_status_label = QLabel("Idle")
        progress_layout.addWidget(self.progress_status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")

        progress_layout.addWidget(self.progress_bar)

        right_layout.addWidget(progress_group)

        # RUN BUTTON
        run_row = QHBoxLayout()
        run_row.addStretch(1)

        self.btn_run = QPushButton("Run Search / Replace")
        run_row.addWidget(self.btn_run)

        right_layout.addLayout(run_row)
        right_layout.addStretch(1)

        right_scroll.setWidget(right_container)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([320, 980])

        # SIGNALS
        self.btn_clear_log.clicked.connect(self.log_box.clear)
        self.btn_browse_output.clicked.connect(self._browse_output_folder)

        self.btn_add_ppts.clicked.connect(self._browse_pptx_files)
        self.btn_remove_ppts.clicked.connect(self._remove_selected_ppt_rows)
        self.btn_clear_ppts.clicked.connect(lambda: self.ppt_table.setRowCount(0))

        self.btn_add_pair.clicked.connect(self._add_pair_row)
        self.btn_remove_pair.clicked.connect(self._remove_pair_row)
        self.btn_clear_pairs.clicked.connect(lambda: self.pair_table.setRowCount(0))

        self.btn_run.clicked.connect(self._run_search_replace)

    def _prefill_defaults(self) -> None:
        self.output_folder_edit.setText(str(self.default_output_folder))

        if self.pair_table.rowCount() == 0:
            self._add_pair_row("March 30, 2026", "March 31, 2026")

        self._set_progress_idle()

    # =====================================================
    # LOGGING
    # =====================================================
    def _log(self, text: str) -> None:
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    # =====================================================
    # PROGRESS
    # =====================================================
    def _set_progress_idle(self) -> None:
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.progress_status_label.setText("Idle")

    def _set_progress_busy(self, text: str) -> None:
        self.progress_bar.setRange(0, 0)
        self.progress_status_label.setText(text)

    def _set_progress_value(self, value: int, status_text: str | None = None) -> None:
        clamped = max(0, min(100, int(value)))
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(clamped)
        self.progress_bar.setFormat("%p%")
        self.progress_status_label.setText(status_text or f"{clamped}% complete")

    def _extract_progress_value(self, payload) -> int | None:
        if isinstance(payload, (int, float)):
            return max(0, min(100, int(payload)))

        if isinstance(payload, str):
            text = payload.strip()

            match = re.search(r"Progress:\s*(\d{1,3})\s*%", text)
            if match:
                return max(0, min(100, int(match.group(1))))

        return None

    def _on_progress(self, payload) -> None:
        value = self._extract_progress_value(payload)

        if value is not None:
            self._set_progress_value(value)

        if isinstance(payload, str):
            self._log(payload)
        else:
            self._log(f"[SearchReplace] Progress update: {payload}")

    # =====================================================
    # BROWSE / TABLE HELPERS
    # =====================================================
    def _browse_output_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select output folder",
            self.output_folder_edit.text().strip() or str(Path.cwd()),
        )

        if folder:
            self.output_folder_edit.setText(folder)
            self.default_output_folder = Path(folder)

            if getattr(self, "state", None) is not None:
                setattr(self.state, "tab6_output_folder", self.default_output_folder)

    def _browse_pptx_files(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PPTX files",
            str(Path.cwd()),
            "PowerPoint Files (*.pptx)",
        )

        for file_path in file_paths:
            self._add_ppt_row(Path(file_path))

    def _add_ppt_row(self, input_pptx: Path, output_name: str | None = None) -> None:
        input_pptx = Path(input_pptx)

        if output_name is None:
            output_name = f"{input_pptx.stem}.pptx"

        existing = set()

        for row in range(self.ppt_table.rowCount()):
            item = self.ppt_table.item(row, 1)
            if item:
                existing.add(item.text().strip())

        if str(input_pptx) in existing:
            return

        row = self.ppt_table.rowCount()
        self.ppt_table.insertRow(row)

        run_item = QTableWidgetItem()
        run_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsUserCheckable
        )
        run_item.setCheckState(Qt.CheckState.Checked)

        input_item = QTableWidgetItem(str(input_pptx))
        input_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
        )

        output_item = QTableWidgetItem(output_name)
        output_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
        )

        self.ppt_table.setItem(row, 0, run_item)
        self.ppt_table.setItem(row, 1, input_item)
        self.ppt_table.setItem(row, 2, output_item)

    def _remove_selected_ppt_rows(self) -> None:
        selected_rows = sorted(
            {index.row() for index in self.ppt_table.selectedIndexes()},
            reverse=True,
        )

        if not selected_rows:
            QMessageBox.information(
                self,
                "No Selection",
                "Please select one or more PPTX rows to remove.",
            )
            return

        for row in selected_rows:
            self.ppt_table.removeRow(row)

    def _add_pair_row(self, search_text: str = "", replace_text: str = "") -> None:
        row = self.pair_table.rowCount()
        self.pair_table.insertRow(row)
        self.pair_table.setItem(row, 0, QTableWidgetItem(search_text))
        self.pair_table.setItem(row, 1, QTableWidgetItem(replace_text))

    def _remove_pair_row(self) -> None:
        row = self.pair_table.currentRow()

        if row < 0:
            QMessageBox.information(
                self,
                "No Selection",
                "Please select a search/replace pair row to remove.",
            )
            return

        self.pair_table.removeRow(row)

    # =====================================================
    # VALIDATION
    # =====================================================
    def _get_pairs(self) -> list[tuple[str, str]]:
        pairs: list[tuple[str, str]] = []

        for row in range(self.pair_table.rowCount()):
            search_item = self.pair_table.item(row, 0)
            replace_item = self.pair_table.item(row, 1)

            search_text = search_item.text().strip() if search_item else ""
            replace_text = replace_item.text() if replace_item else ""

            if not search_text and not replace_text:
                continue

            if not search_text:
                raise ValueError(f"Row {row + 1}: Search Text is required.")

            pairs.append((search_text, replace_text))

        if not pairs:
            raise ValueError("Please add at least one Search/Replace pair.")

        return pairs

    def _get_jobs(self, output_folder: Path) -> list[tuple[Path, Path]]:
        if self.ppt_table.rowCount() == 0:
            raise ValueError("Please add at least one PPTX file.")

        checked_rows: list[int] = []

        for row in range(self.ppt_table.rowCount()):
            run_item = self.ppt_table.item(row, 0)

            if run_item and run_item.checkState() == Qt.CheckState.Checked:
                checked_rows.append(row)

        rows_to_process = checked_rows or list(range(self.ppt_table.rowCount()))

        jobs: list[tuple[Path, Path]] = []

        for row in rows_to_process:
            input_item = self.ppt_table.item(row, 1)
            output_item = self.ppt_table.item(row, 2)

            input_text = input_item.text().strip() if input_item else ""
            output_name = output_item.text().strip() if output_item else ""

            if not input_text:
                raise ValueError(f"Row {row + 1}: Input PPTX is required.")

            input_pptx = Path(input_text)

            if not input_pptx.exists():
                raise ValueError(f"Row {row + 1}: Input PPTX does not exist:\n{input_pptx}")

            if not output_name:
                raise ValueError(f"Row {row + 1}: Output PPTX name is required.")

            if not output_name.lower().endswith(".pptx"):
                output_name += ".pptx"

            output_pptx = output_folder / output_name

            if input_pptx.resolve() == output_pptx.resolve():
                raise ValueError(
                    f"Row {row + 1}: Output PPTX cannot be same as input PPTX.\n"
                    f"Please use a different output name."
                )

            jobs.append((input_pptx, output_pptx))

        return jobs

    def _validate_inputs(
        self,
    ) -> tuple[list[tuple[Path, Path]], list[tuple[str, str]], list[str], bool] | None:
        output_folder_text = self.output_folder_edit.text().strip()

        if not output_folder_text:
            QMessageBox.warning(self, "Invalid Output", "Please select a valid output folder.")
            return None

        output_folder = Path(output_folder_text)

        try:
            output_folder.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Invalid Output",
                f"Could not create output folder:\n{output_folder}\n\n{exc}",
            )
            return None

        try:
            jobs = self._get_jobs(output_folder)
            pairs = self._get_pairs()
        except Exception as exc:
            QMessageBox.warning(self, "Invalid Search/Replace Input", str(exc))
            return None

        blacklist = [
            part.strip()
            for part in self.blacklist_edit.text().split(",")
            if part.strip()
        ]

        case_sensitive = self.case_sensitive_chk.isChecked()

        return jobs, pairs, blacklist, case_sensitive

    # =====================================================
    # RUN
    # =====================================================
    def _run_search_replace(self) -> None:
        validated = self._validate_inputs()

        if validated is None:
            return

        jobs, pairs, blacklist, case_sensitive = validated

        self._log("=" * 70)
        self._log(f"[SearchReplace] PPTX files: {len(jobs)}")
        self._log(f"[SearchReplace] Pairs: {pairs}")
        self._log(f"[SearchReplace] Case sensitive: {case_sensitive}")
        self._log(f"[SearchReplace] Blacklist: {blacklist}")
        self._log("[SearchReplace] Starting batch search/replace...")

        self.btn_run.setEnabled(False)
        self._set_progress_busy("Running search/replace...")

        worker = FunctionWorker(
            run_search_replace_batch_action,
            jobs,
            pairs,
            blacklist,
            case_sensitive,
        )

        self.worker = worker

        worker.signals.started.connect(lambda: self._log("[SearchReplace] Worker started"))
        worker.signals.progress.connect(self._on_progress)
        worker.signals.finished.connect(self._on_finished)
        worker.signals.error.connect(self._on_error)

        self.thread_pool.start(worker)

    def _on_finished(self, result) -> None:
        try:
            self.worker = None
            self.btn_run.setEnabled(True)

            output_paths = [Path(path) for path in result]

            self._set_progress_value(100, "Completed")
            self._log("[SearchReplace] Finished successfully.")

            for output_path in output_paths:
                self._log(f"[SearchReplace] Saved: {output_path}")

            if getattr(self, "state", None) is not None and output_paths:
                setattr(self.state, "tab6_output_folder", output_paths[0].parent)

            QMessageBox.information(
                self,
                "Search/Replace Complete",
                "Updated PPTX file(s) saved to:\n\n"
                + "\n".join(str(path) for path in output_paths),
            )

        except Exception:
            error_text = traceback.format_exc()

            self.worker = None
            self.btn_run.setEnabled(True)
            self._set_progress_idle()
            self.progress_status_label.setText("Failed")

            self._log("[SearchReplace] ERROR in finish handler")
            self._log(error_text)

            QMessageBox.critical(self, "Search/Replace Finish Failed", error_text)

    def _on_error(self, error_text: str) -> None:
        self.worker = None
        self.btn_run.setEnabled(True)

        self._set_progress_idle()
        self.progress_status_label.setText("Failed")

        self._log("[SearchReplace] ERROR")
        self._log(error_text)

        QMessageBox.critical(self, "Search/Replace Failed", error_text)


Tab6Widget = SearchReplaceWidget
SearchAndReplaceWidget = SearchReplaceWidget


# =========================================================
# DIRECT RUNNER
# =========================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = SearchReplaceWidget()
    window.resize(1350, 750)
    window.show()

    sys.exit(app.exec())