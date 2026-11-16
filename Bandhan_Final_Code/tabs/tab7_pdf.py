from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Iterable
import re
import shutil
import sys
import traceback

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
    QAbstractItemView,
    QScrollArea,
    QSplitter,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
)


FULL_PATH_ROLE = Qt.ItemDataRole.UserRole
PAGE_COUNT_ROLE = Qt.ItemDataRole.UserRole.value + 1
PP_SAVE_AS_PDF = 32


# =========================================================
# WORKER
# =========================================================
class WorkerSignals(QObject):
    started = pyqtSignal()
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)


class FunctionWorker(QRunnable):
    """Run a plain Python callable in Qt thread pool."""

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
# BASIC HELPERS
# =========================================================
def natural_key(text: str):
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", str(text))
    ]


def get_pdf_page_count(pdf_path: str | Path) -> int:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("pypdf is required. Install with: pip install pypdf") from exc

    reader = PdfReader(str(pdf_path))
    return len(reader.pages)


def count_slides(pptx_path: str | Path) -> int:
    """Count slides without PowerPoint."""
    from pptx import Presentation

    prs = Presentation(str(Path(pptx_path)))
    return len(prs.slides)


# =========================================================
# POWERPOINT COM HELPERS
# =========================================================
def _ensure_windows() -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("PowerPoint COM automation only works on Windows.")


def _start_powerpoint(visible: bool = False):
    """
    Start PowerPoint COM safely.

    Important:
    Do not set app.Visible = 0 / False.
    Some PowerPoint versions throw:
    'Application.Visible : Invalid request. Hiding the application window is not allowed.'
    """
    _ensure_windows()

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for PowerPoint COM automation. "
            "Install with: pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()

    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")

        if visible:
            try:
                app.Visible = 1
            except Exception:
                pass

        try:
            app.DisplayAlerts = 0
        except Exception:
            pass

        return app, pythoncom

    except Exception:
        pythoncom.CoUninitialize()
        raise


def _stop_powerpoint(app, pythoncom_module) -> None:
    try:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
    finally:
        if pythoncom_module is not None:
            try:
                pythoncom_module.CoUninitialize()
            except Exception:
                pass


def _export_pptx_to_pdf_with_app(app, pptx_path: str | Path, pdf_path: str | Path) -> Path:
    """
    Export one PPTX to PDF using an already-open PowerPoint COM application.
    """
    pptx_path = Path(pptx_path).resolve()
    pdf_path = Path(pdf_path).resolve()

    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX file not found on disk: {pptx_path}")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    presentation = None

    try:
        presentation = app.Presentations.Open(
            FileName=str(pptx_path),
            ReadOnly=False,
            Untitled=False,
            WithWindow=False,
        )

        presentation.SaveAs(str(pdf_path), PP_SAVE_AS_PDF)

    except Exception as exc:
        raise RuntimeError(f"Failed to export PPTX to PDF:\n{pptx_path}\n\n{exc}") from exc

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass

    return pdf_path


def export_pptx_to_pdf(pptx_path: str | Path, pdf_path: str | Path) -> Path:
    """
    Export one PPTX to PDF.

    For converting many files, use convert_folder_pptx_to_pdf(),
    because that opens PowerPoint only once.
    """
    app = None
    pythoncom_module = None

    try:
        app, pythoncom_module = _start_powerpoint(visible=False)
        return _export_pptx_to_pdf_with_app(app, pptx_path, pdf_path)

    finally:
        _stop_powerpoint(app, pythoncom_module)


def merge_presentations(source_pptxs: Iterable[str | Path], output_pptx: str | Path) -> Path:
    """
    Optional utility:
    Merge complete PPTX presentations in order.
    """
    source_pptxs = [Path(p).resolve() for p in source_pptxs]

    if not source_pptxs:
        raise ValueError("No source presentations provided.")

    for pptx_path in source_pptxs:
        if not pptx_path.exists():
            raise FileNotFoundError(f"PPTX file not found: {pptx_path}")

    output_pptx = Path(output_pptx).resolve()
    output_pptx.parent.mkdir(parents=True, exist_ok=True)

    app = None
    pythoncom_module = None
    target = None

    try:
        app, pythoncom_module = _start_powerpoint(visible=False)

        shutil.copyfile(source_pptxs[0], output_pptx)

        target = app.Presentations.Open(
            FileName=str(output_pptx),
            ReadOnly=False,
            Untitled=False,
            WithWindow=False,
        )

        for src_path in source_pptxs[1:]:
            target.Slides.InsertFromFile(str(src_path), target.Slides.Count)

        target.Save()

    finally:
        if target is not None:
            try:
                target.Close()
            except Exception:
                pass

        _stop_powerpoint(app, pythoncom_module)

    return output_pptx


# =========================================================
# PPTX TO PDF / PDF MERGE LOGIC
# =========================================================
def convert_folder_pptx_to_pdf(
    input_folder: str | Path,
    output_folder: str | Path,
    progress_cb: Callable[[str], None] | None = None,
) -> list[Path]:
    """
    Convert all PPTX files in input_folder to PDF into output_folder.
    Returns successfully created PDF paths.
    """

    def log(msg: str) -> None:
        if progress_cb:
            progress_cb(msg)

    input_folder = Path(input_folder)
    output_folder = Path(output_folder)

    if not input_folder.exists() or not input_folder.is_dir():
        raise ValueError(f"Invalid input folder: {input_folder}")

    pptx_files = sorted(
        [
            p
            for p in input_folder.glob("*.pptx")
            if p.is_file() and not p.name.startswith("~$")
        ],
        key=lambda p: natural_key(p.name),
    )

    if not pptx_files:
        raise ValueError(f"No .pptx files found in: {input_folder}")

    output_folder.mkdir(parents=True, exist_ok=True)

    total_files = len(pptx_files)
    processed_files = 0
    last_percent = -1

    pdf_paths: list[Path] = []
    failed: list[str] = []

    log(f"[Tab6] Found {total_files} PPTX file(s) for conversion")
    log("[Tab6] Progress: 0%")

    app = None
    pythoncom_module = None

    try:
        app, pythoncom_module = _start_powerpoint(visible=False)

        for pptx_path in pptx_files:
            pdf_path = output_folder / f"{pptx_path.stem}.pdf"

            log(f"[Tab6] Opening PPTX: {pptx_path}")

            try:
                _export_pptx_to_pdf_with_app(app, pptx_path, pdf_path)
                pdf_paths.append(pdf_path)
                log(f"[Tab6] PDF created: {pdf_path}")

            except Exception as exc:
                msg = f"Failed converting '{pptx_path.name}': {exc}"
                failed.append(msg)
                log(f"[Tab6] ERROR: {msg}")

            processed_files += 1
            percent = int((processed_files / total_files) * 100)

            if percent != last_percent:
                log(f"[Tab6] Progress: {percent}%")
                last_percent = percent

    finally:
        _stop_powerpoint(app, pythoncom_module)

    if failed:
        log("=" * 70)
        log("[Tab6] Some files failed during conversion:")
        for item in failed:
            log(f"  - {item}")

    if not pdf_paths:
        raise RuntimeError("No PDFs were created successfully.")

    if last_percent < 100:
        log("[Tab6] Progress: 100%")

    return pdf_paths


def merge_pdf_files(
    pdf_paths: list[str | Path],
    output_pdf: str | Path,
    progress_cb: Callable[[str], None] | None = None,
) -> Path:
    """
    Merge PDFs in given order.
    Writes to a temp PDF first, then replaces final output safely.
    """

    def log(msg: str) -> None:
        if progress_cb:
            progress_cb(msg)

    try:
        from pypdf import PdfWriter
    except ImportError as exc:
        raise RuntimeError("pypdf is required. Install with: pip install pypdf") from exc

    pdf_paths = [Path(p).resolve() for p in pdf_paths]

    if not pdf_paths:
        raise ValueError("No PDF paths provided for merging.")

    for pdf_path in pdf_paths:
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    output_pdf = Path(output_pdf).resolve()
    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    if output_pdf in pdf_paths:
        raise ValueError(
            "Output PDF is also selected as an input PDF. "
            "Please choose a different final PDF name."
        )

    temp_pdf = output_pdf.with_name(output_pdf.stem + "_temp_merge.pdf")

    total_files = len(pdf_paths)
    processed_files = 0
    last_percent = -1

    writer = PdfWriter()

    try:
        log(f"[Tab6] Merging {total_files} PDF file(s)")
        log("[Tab6] Progress: 0%")

        for pdf_path in pdf_paths:
            writer.append(str(pdf_path))

            processed_files += 1
            percent = int((processed_files / total_files) * 100)

            log(f"[Tab6] Added: {pdf_path}")

            if percent != last_percent:
                log(f"[Tab6] Progress: {percent}%")
                last_percent = percent

        if temp_pdf.exists():
            temp_pdf.unlink()

        with open(temp_pdf, "wb") as f:
            writer.write(f)

    finally:
        try:
            writer.close()
        except Exception:
            pass

    if output_pdf.exists():
        output_pdf.unlink()

    temp_pdf.replace(output_pdf)

    if last_percent < 100:
        log("[Tab6] Progress: 100%")

    log(f"[Tab6] Final merged PDF: {output_pdf}")

    return output_pdf


# =========================================================
# UI WIDGET
# =========================================================
class PdfWidget(QWidget):
    """
    Convert PPTX folder to PDF, load PDFs, reorder them,
    and merge checked PDFs only.

    state is optional.
    If state is provided, these attributes are used if available:
        state.tab5_updated_folder
        state.tab8_pdf_folder
        state.tab8_final_pdf
        state.ensure_output_subdir("08_pdf")
    """

    COL_MERGE = 0
    COL_FILE = 1
    COL_PAGES = 2

    def __init__(self, state: Any | None = None):
        super().__init__()

        self.state = state
        self.thread_pool = QThreadPool.globalInstance()

        self.convert_worker: FunctionWorker | None = None
        self.merge_worker: FunctionWorker | None = None

        self.click_order_mode = False
        self.click_order_sequence: list[int] = []

        self._build_ui()
        self._prefill_from_state()

    # -----------------------------------------------------
    # State helpers
    # -----------------------------------------------------
    def _state_get(self, name: str, default=None):
        return getattr(self.state, name, default) if self.state is not None else default

    def _state_set(self, name: str, value) -> None:
        if self.state is not None:
            setattr(self.state, name, value)

    def _default_pdf_folder(self) -> Path:
        if self.state is not None and hasattr(self.state, "ensure_output_subdir"):
            try:
                return Path(self.state.ensure_output_subdir("08_pdf"))
            except Exception:
                pass

        return self._app_run_folder() / "output" / "08_pdf"

    # -----------------------------------------------------
    # Build UI
    # -----------------------------------------------------
    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(8)

        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter)

        # LEFT SIDE
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(8)
        left_layout.setContentsMargins(0, 0, 0, 0)

        title = QLabel("Convert PPTX to PDF and Merge")
        title.setStyleSheet("font-weight: bold; font-size: 14px;")
        left_layout.addWidget(title)

        info = QLabel(
            "Convert PPTX files to PDF, load/reorder PDFs, and merge checked PDFs only.\n"
            "The PDF table shows file name and page count."
        )
        info.setWordWrap(True)
        left_layout.addWidget(info)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("PDF conversion and merge logs will appear here...")
        left_layout.addWidget(self.log_box, stretch=1)

        self.btn_clear_log = QPushButton("Clear Log")
        left_layout.addWidget(self.btn_clear_log)

        # RIGHT SIDE
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)

        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setSpacing(8)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # Conversion inputs
        group = QGroupBox("PDF Conversion")
        grid = QGridLayout(group)
        grid.setContentsMargins(8, 8, 8, 8)
        grid.setSpacing(6)

        self.input_folder_edit = QLineEdit()
        self.btn_browse_input = QPushButton("Browse...")

        grid.addWidget(QLabel("Input PPTX Folder:"), 0, 0)
        grid.addWidget(self.input_folder_edit, 0, 1)
        grid.addWidget(self.btn_browse_input, 0, 2)

        self.pdf_folder_edit = QLineEdit()
        self.btn_browse_pdf_folder = QPushButton("Browse...")

        grid.addWidget(QLabel("PDF Output Folder:"), 1, 0)
        grid.addWidget(self.pdf_folder_edit, 1, 1)
        grid.addWidget(self.btn_browse_pdf_folder, 1, 2)

        self.final_pdf_edit = QLineEdit("Bandhan_Final_Factsheet.pdf")

        grid.addWidget(QLabel("Final Merged PDF Name:"), 2, 0)
        grid.addWidget(self.final_pdf_edit, 2, 1)

        right_layout.addWidget(group)

        # Progress
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setContentsMargins(8, 8, 8, 8)
        progress_layout.setSpacing(6)

        self.progress_status_label = QLabel("Idle")
        progress_layout.addWidget(self.progress_status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setTextVisible(True)
        progress_layout.addWidget(self.progress_bar)

        right_layout.addWidget(progress_group)

        # Buttons
        row1 = QHBoxLayout()
        row1.setSpacing(6)

        self.btn_convert = QPushButton("Convert All PPTX to PDF")
        self.btn_load_pdfs_from_folder = QPushButton("Load PDFs From Folder")
        self.btn_add_pdf_files = QPushButton("Add PDF Files...")

        row1.addWidget(self.btn_convert)
        row1.addWidget(self.btn_load_pdfs_from_folder)
        row1.addWidget(self.btn_add_pdf_files)
        row1.addStretch(1)

        right_layout.addLayout(row1)

        # Table
        table_group = QGroupBox("PDFs to Merge / Reorder")
        table_layout = QVBoxLayout(table_group)
        table_layout.setContentsMargins(8, 8, 8, 8)
        table_layout.setSpacing(6)

        table_help = QLabel(
            "Only checked PDFs will be merged. "
            "Use Move Up / Move Down or Click Order Mode to set merge order."
        )
        table_help.setWordWrap(True)
        table_layout.addWidget(table_help)

        self.checked_pages_label = QLabel("Checked PDFs: 0 | Total pages after merge: 0")
        self.checked_pages_label.setStyleSheet("font-weight: bold;")
        table_layout.addWidget(self.checked_pages_label)

        self.pdf_table = QTableWidget()
        self.pdf_table.setColumnCount(3)
        self.pdf_table.setHorizontalHeaderLabels(["Merge", "File Name", "Pages"])
        self.pdf_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.pdf_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.pdf_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.pdf_table.setMinimumHeight(360)
        self.pdf_table.verticalHeader().setVisible(False)

        header = self.pdf_table.horizontalHeader()
        header.setSectionResizeMode(self.COL_MERGE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_FILE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_PAGES, QHeaderView.ResizeMode.ResizeToContents)

        table_layout.addWidget(self.pdf_table, stretch=1)

        reorder_row = QHBoxLayout()
        reorder_row.setSpacing(6)

        self.btn_click_order = QPushButton("Click Order Mode")
        self.btn_up = QPushButton("Move Up")
        self.btn_down = QPushButton("Move Down")
        self.btn_remove_selected = QPushButton("Remove Selected PDFs")
        self.btn_clear_pdf_list = QPushButton("Clear PDF Table")

        reorder_row.addWidget(self.btn_click_order)
        reorder_row.addWidget(self.btn_up)
        reorder_row.addWidget(self.btn_down)
        reorder_row.addWidget(self.btn_remove_selected)
        reorder_row.addWidget(self.btn_clear_pdf_list)
        reorder_row.addStretch(1)

        table_layout.addLayout(reorder_row)
        right_layout.addWidget(table_group, stretch=1)

        row2 = QHBoxLayout()
        row2.addStretch(1)

        self.btn_merge = QPushButton("Merge Checked PDFs")
        row2.addWidget(self.btn_merge)

        right_layout.addLayout(row2)

        right_scroll.setWidget(right_container)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([360, 980])

        # Signals
        self.btn_browse_input.clicked.connect(self._browse_input)
        self.btn_browse_pdf_folder.clicked.connect(self._browse_pdf_folder)

        self.btn_convert.clicked.connect(self._run_convert)
        self.btn_load_pdfs_from_folder.clicked.connect(self._load_pdfs_from_folder)
        self.btn_add_pdf_files.clicked.connect(self._add_pdf_files)

        self.btn_click_order.clicked.connect(self._toggle_click_order_mode)
        self.pdf_table.cellClicked.connect(self._handle_click_order_cell)
        self.pdf_table.itemChanged.connect(self._refresh_checked_pages_total)

        self.btn_up.clicked.connect(self._move_up)
        self.btn_down.clicked.connect(self._move_down)
        self.btn_remove_selected.clicked.connect(self._remove_selected_pdfs)
        self.btn_clear_pdf_list.clicked.connect(self._clear_pdf_table)

        self.btn_merge.clicked.connect(self._run_merge)
        self.btn_clear_log.clicked.connect(self.log_box.clear)

    # -----------------------------------------------------
    # Prefill
    # -----------------------------------------------------
    def _prefill_from_state(self):
        tab5_folder = self._state_get("tab7_output_folder") or self._state_get("tab5_updated_folder")

        if tab5_folder:
            self.input_folder_edit.setText(str(tab5_folder))

        if not self.pdf_folder_edit.text().strip():
            self.pdf_folder_edit.setText(str(self._default_pdf_folder()))

        self._set_progress_idle()
        self._refresh_checked_pages_total()

    # -----------------------------------------------------
    # Browse helpers
    # -----------------------------------------------------
    def _browse_input(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select input PPTX folder",
            self.input_folder_edit.text().strip() or str(Path.cwd()),
        )

        if folder:
            self.input_folder_edit.setText(folder)


    def _browse_pdf_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select PDF output folder",
            self.pdf_folder_edit.text().strip() or str(Path.cwd()),
        )

        if folder:
            self.pdf_folder_edit.setText(folder)

    def _log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    # -----------------------------------------------------
    # Progress helpers
    # -----------------------------------------------------
    def _set_progress_idle(self):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.progress_status_label.setText("Idle")

    def _set_progress_busy(self, status_text: str = "Working..."):
        self.progress_bar.setRange(0, 0)
        self.progress_status_label.setText(status_text)

    def _set_progress_value(self, value: int, status_text: str | None = None):
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

            if text.isdigit():
                return max(0, min(100, int(text)))

            match = re.search(r"(\d{1,3})\s*%", text)
            if match:
                return max(0, min(100, int(match.group(1))))

        return None

    def _on_progress(self, payload):
        progress_value = self._extract_progress_value(payload)

        if progress_value is not None:
            self._set_progress_value(progress_value)

        if isinstance(payload, str):
            self._log(payload)
        else:
            self._log(f"[Tab6] Progress update: {payload}")

    # -----------------------------------------------------
    # Table helpers
    # -----------------------------------------------------
    def _row_count(self) -> int:
        return self.pdf_table.rowCount()

    def _file_item(self, row: int) -> QTableWidgetItem | None:
        return self.pdf_table.item(row, self.COL_FILE)

    def _merge_item(self, row: int) -> QTableWidgetItem | None:
        return self.pdf_table.item(row, self.COL_MERGE)

    def _pdf_row_path(self, row: int) -> str:
        item = self._file_item(row)
        return str(item.data(FULL_PATH_ROLE) if item else "").strip()

    def _pdf_row_pages(self, row: int) -> int:
        item = self._file_item(row)

        try:
            return int(item.data(PAGE_COUNT_ROLE) if item else 0)
        except Exception:
            return 0

    def _is_row_checked(self, row: int) -> bool:
        item = self._merge_item(row)
        return bool(item and item.checkState() == Qt.CheckState.Checked)

    def _row_data(self, row: int) -> dict:
        return {
            "path": self._pdf_row_path(row),
            "pages": self._pdf_row_pages(row),
            "checked": self._is_row_checked(row),
        }

    def _set_row_data(self, row: int, data: dict, prefix: str = ""):
        pdf_path = Path(data["path"])
        page_count = int(data["pages"])
        checked = bool(data["checked"])

        merge_item = QTableWidgetItem()
        merge_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsUserCheckable
        )
        merge_item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        merge_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        file_item = QTableWidgetItem(f"{prefix}{pdf_path.name}")
        file_item.setData(FULL_PATH_ROLE, str(pdf_path))
        file_item.setData(PAGE_COUNT_ROLE, page_count)

        pages_item = QTableWidgetItem(str(page_count))
        pages_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        self.pdf_table.setItem(row, self.COL_MERGE, merge_item)
        self.pdf_table.setItem(row, self.COL_FILE, file_item)
        self.pdf_table.setItem(row, self.COL_PAGES, pages_item)

    def _refresh_table_row_labels(self):
        order_map = {
            row: index + 1
            for index, row in enumerate(self.click_order_sequence)
        }

        self.pdf_table.blockSignals(True)

        try:
            for row in range(self._row_count()):
                file_item = self._file_item(row)

                if not file_item:
                    continue

                pdf_path = Path(self._pdf_row_path(row))
                prefix = f"{order_map[row]}. " if row in order_map else ""
                file_item.setText(f"{prefix}{pdf_path.name}")

        finally:
            self.pdf_table.blockSignals(False)

    def _refresh_checked_pages_total(self, *args):
        checked_count = 0
        total_pages = 0

        for row in range(self._row_count()):
            if self._is_row_checked(row):
                checked_count += 1
                total_pages += self._pdf_row_pages(row)

        self.checked_pages_label.setText(
            f"Checked PDFs: {checked_count} | Total pages after merge: {total_pages}"
        )

    def _add_pdf_to_table(self, pdf_path: str | Path):
        pdf_path = Path(pdf_path).resolve()

        existing_paths = {
            self._pdf_row_path(row)
            for row in range(self._row_count())
        }

        if str(pdf_path) in existing_paths:
            return

        try:
            page_count = get_pdf_page_count(pdf_path)
        except Exception as exc:
            page_count = 0
            self._log(f"[Tab6] WARNING: Could not read page count for {pdf_path.name}: {exc}")

        row = self._row_count()

        self.pdf_table.blockSignals(True)

        try:
            self.pdf_table.insertRow(row)
            self._set_row_data(
                row,
                {
                    "path": str(pdf_path),
                    "pages": page_count,
                    "checked": True,
                },
            )

        finally:
            self.pdf_table.blockSignals(False)

        self._refresh_checked_pages_total()

    def _load_pdfs_from_folder(self):
        self._cancel_click_order_if_active()

        folder_text = self.pdf_folder_edit.text().strip()

        if not folder_text:
            QMessageBox.warning(self, "No PDF Folder", "Please select a PDF folder first.")
            return

        pdf_folder = Path(folder_text)

        if not pdf_folder.exists() or not pdf_folder.is_dir():
            QMessageBox.warning(self, "Invalid PDF Folder", "Please select a valid PDF folder.")
            return

        pdf_files = sorted(pdf_folder.glob("*.pdf"), key=lambda p: natural_key(p.name))

        if not pdf_files:
            QMessageBox.information(self, "No PDFs", f"No PDF files found in:\n{pdf_folder}")
            return

        self.pdf_table.setRowCount(0)

        for pdf_path in pdf_files:
            self._add_pdf_to_table(pdf_path)

        self._refresh_checked_pages_total()
        self._log(f"[Tab6] Loaded {len(pdf_files)} PDFs from folder: {pdf_folder}")

    def _add_pdf_files(self):
        self._cancel_click_order_if_active()

        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF files to add",
            self.pdf_folder_edit.text().strip() or str(Path.cwd()),
            "PDF Files (*.pdf)",
        )

        if not file_paths:
            return

        added = 0

        for file_path in file_paths:
            before = self._row_count()
            self._add_pdf_to_table(Path(file_path))

            if self._row_count() > before:
                added += 1

        self._refresh_checked_pages_total()
        self._log(f"[Tab6] Added {added} PDF files manually.")

    def _remove_selected_pdfs(self):
        self._cancel_click_order_if_active()

        rows = sorted(
            {index.row() for index in self.pdf_table.selectionModel().selectedRows()},
            reverse=True,
        )

        if not rows:
            QMessageBox.information(self, "No Selection", "Please select one or more PDFs to remove.")
            return

        for row in rows:
            self.pdf_table.removeRow(row)

        self._refresh_checked_pages_total()

    def _clear_pdf_table(self):
        self._cancel_click_order_if_active()
        self.pdf_table.setRowCount(0)
        self._refresh_checked_pages_total()

    def _move_up(self):
        self._cancel_click_order_if_active()

        row = self.pdf_table.currentRow()

        if row <= 0:
            return

        current_data = self._row_data(row)
        previous_data = self._row_data(row - 1)

        self.pdf_table.blockSignals(True)

        try:
            self._set_row_data(row - 1, current_data)
            self._set_row_data(row, previous_data)

        finally:
            self.pdf_table.blockSignals(False)

        self.pdf_table.selectRow(row - 1)
        self._refresh_checked_pages_total()

    def _move_down(self):
        self._cancel_click_order_if_active()

        row = self.pdf_table.currentRow()

        if row < 0 or row >= self._row_count() - 1:
            return

        current_data = self._row_data(row)
        next_data = self._row_data(row + 1)

        self.pdf_table.blockSignals(True)

        try:
            self._set_row_data(row + 1, current_data)
            self._set_row_data(row, next_data)

        finally:
            self.pdf_table.blockSignals(False)

        self.pdf_table.selectRow(row + 1)
        self._refresh_checked_pages_total()

    def _get_pdf_paths_for_merge(self) -> list[Path]:
        result: list[Path] = []

        for row in range(self._row_count()):
            if self._is_row_checked(row):
                result.append(Path(self._pdf_row_path(row)))

        return result

    def _get_checked_pages_total(self) -> int:
        total_pages = 0

        for row in range(self._row_count()):
            if self._is_row_checked(row):
                total_pages += self._pdf_row_pages(row)

        return total_pages

    # -----------------------------------------------------
    # Click order helpers
    # -----------------------------------------------------
    def _toggle_click_order_mode(self):
        if self._row_count() == 0:
            QMessageBox.information(self, "No PDFs", "Load or add PDFs first.")
            return

        self.click_order_mode = not self.click_order_mode

        if self.click_order_mode:
            self.pdf_table.clearSelection()
            self.click_order_sequence = []
            self.btn_click_order.setText("Cancel Click Order")
            self._refresh_table_row_labels()
            self._log("[Tab6] Click Order Mode enabled. Click PDF rows in required order.")

            QMessageBox.information(
                self,
                "Click Order Mode",
                "Click PDF rows one by one in the desired merge order.\n\n"
                "Click an already selected row again to remove it from the order.\n"
                "When all rows are selected, the table will reorder automatically.\n\n"
                "Only checked PDFs will be merged.",
            )
        else:
            self._clear_click_order_marks()
            self.btn_click_order.setText("Click Order Mode")
            self._log("[Tab6] Click Order Mode cancelled.")

    def _handle_click_order_cell(self, row: int, column: int):
        if not self.click_order_mode:
            return

        if row < 0:
            return

        pdf_name = Path(self._pdf_row_path(row)).name

        if row in self.click_order_sequence:
            self.click_order_sequence.remove(row)
            self._log(f"[Tab6] Removed from click order: {pdf_name}")
        else:
            self.click_order_sequence.append(row)
            self._log(
                f"[Tab6] Click order selected "
                f"{len(self.click_order_sequence)}/{self._row_count()}: {pdf_name}"
            )

        self._refresh_table_row_labels()

        if len(self.click_order_sequence) == self._row_count():
            self._apply_click_order()

    def _clear_click_order_marks(self):
        self.click_order_sequence = []
        self._refresh_table_row_labels()

    def _cancel_click_order_if_active(self):
        if self.click_order_mode:
            self.click_order_mode = False
            self._clear_click_order_marks()
            self.btn_click_order.setText("Click Order Mode")
            self._log("[Tab6] Click Order Mode cancelled.")

    def _apply_click_order(self):
        self._log("[Tab6] Applying clicked PDF order...")

        ordered_data = [
            self._row_data(row)
            for row in self.click_order_sequence
            if 0 <= row < self._row_count()
        ]

        self.pdf_table.blockSignals(True)

        try:
            self.pdf_table.setRowCount(0)

            for data in ordered_data:
                row = self._row_count()
                self.pdf_table.insertRow(row)
                self._set_row_data(row, data)

        finally:
            self.pdf_table.blockSignals(False)

        self.click_order_mode = False
        self.click_order_sequence = []
        self.btn_click_order.setText("Click Order Mode")

        self._refresh_checked_pages_total()
        self._log("[Tab6] Click ordering applied successfully.")

    # -----------------------------------------------------
    # Conversion
    # -----------------------------------------------------
    def _validate_convert(self):
        input_folder_text = self.input_folder_edit.text().strip()

        if not input_folder_text:
            QMessageBox.warning(self, "Invalid Input Folder", "Please select a valid input folder.")
            return None

        input_folder = Path(input_folder_text)

        if not input_folder.exists() or not input_folder.is_dir():
            QMessageBox.warning(self, "Invalid Input Folder", "Please select a valid input folder.")
            return None

        pdf_folder_text = self.pdf_folder_edit.text().strip()

        if not pdf_folder_text:
            QMessageBox.warning(self, "Invalid PDF Folder", "Please select a valid PDF output folder.")
            return None

        pdf_folder = Path(pdf_folder_text)
        pdf_folder.mkdir(parents=True, exist_ok=True)

        return input_folder, pdf_folder

    def _run_convert(self):
        self._cancel_click_order_if_active()

        validated = self._validate_convert()

        if validated is None:
            return

        input_folder, pdf_folder = validated

        self._log("=" * 70)
        self._log(f"[Tab6] Input PPTX folder: {input_folder}")
        self._log(f"[Tab6] PDF output folder: {pdf_folder}")
        self._log("[Tab6] Starting PDF conversion...")

        self.btn_convert.setEnabled(False)
        self._set_progress_busy("Converting PPTX files to PDF...")

        worker = FunctionWorker(
            convert_folder_pptx_to_pdf,
            input_folder,
            pdf_folder,
        )

        self.convert_worker = worker

        worker.signals.started.connect(lambda: self._log("[Tab6] Convert worker started"))
        worker.signals.progress.connect(self._on_progress)
        worker.signals.finished.connect(self._on_convert_finished)
        worker.signals.error.connect(self._on_convert_error)

        self.thread_pool.start(worker)

    def _on_convert_finished(self, result):
        self.convert_worker = None
        self.btn_convert.setEnabled(True)
        self.pdf_table.setRowCount(0)

        for pdf_path in result:
            self._add_pdf_to_table(Path(pdf_path))

        pdf_folder = Path(self.pdf_folder_edit.text().strip())
        self._state_set("tab8_pdf_folder", pdf_folder)

        total_pages = self._get_checked_pages_total()

        self._set_progress_value(100, "Completed")
        self._refresh_checked_pages_total()

        self._log(f"[Tab6] PDF conversion complete. Files: {len(result)}")
        self._log(f"[Tab6] Total pages after merge if all remain checked: {total_pages}")

        QMessageBox.information(
            self,
            "PDF Conversion Complete",
            f"Converted {len(result)} PDF files.\n"
            f"Total pages selected for merge: {total_pages}\n\n"
            f"PDF folder:\n{pdf_folder}",
        )

    def _on_convert_error(self, error_text: str):
        self.convert_worker = None
        self.btn_convert.setEnabled(True)
        self._set_progress_idle()
        self.progress_status_label.setText("Failed")

        self._log("[Tab6] CONVERT ERROR")
        self._log(error_text)

        QMessageBox.critical(self, "Convert Failed", error_text)

    # -----------------------------------------------------
    # Merge
    # -----------------------------------------------------
    def _run_merge(self):
        self._cancel_click_order_if_active()

        if self._row_count() == 0:
            QMessageBox.warning(
                self,
                "No PDFs",
                "No PDFs available to merge. Convert, load, or add PDFs first.",
            )
            return

        pdf_folder_text = self.pdf_folder_edit.text().strip()

        if not pdf_folder_text:
            QMessageBox.warning(self, "Invalid PDF Folder", "Please select a valid PDF output folder.")
            return

        pdf_folder = Path(pdf_folder_text)
        pdf_folder.mkdir(parents=True, exist_ok=True)

        final_name = self.final_pdf_edit.text().strip() or "ABSL_Final.pdf"

        if not final_name.lower().endswith(".pdf"):
            final_name += ".pdf"

        output_pdf = pdf_folder / final_name
        pdf_paths = self._get_pdf_paths_for_merge()
        total_pages = self._get_checked_pages_total()

        if not pdf_paths:
            QMessageBox.warning(
                self,
                "No PDFs Checked",
                "Please check at least one PDF to include in the merge.",
            )
            return

        if output_pdf.resolve() in [p.resolve() for p in pdf_paths]:
            QMessageBox.warning(
                self,
                "Invalid Output Name",
                "The final output PDF is also selected in the table.\n\n"
                "Please choose a different final PDF name.",
            )
            return

        self._log("=" * 70)
        self._log(f"[Tab6] Merging {len(pdf_paths)} checked PDF(s)...")
        self._log(f"[Tab6] Total pages after merge: {total_pages}")
        self._log(f"[Tab6] Final output: {output_pdf}")

        for pdf_path in pdf_paths:
            self._log(f"[Tab6] Checked for merge: {pdf_path.name}")

        self.btn_merge.setEnabled(False)
        self._set_progress_busy("Merging PDF files...")

        worker = FunctionWorker(
            merge_pdf_files,
            pdf_paths,
            output_pdf,
        )

        self.merge_worker = worker

        worker.signals.started.connect(lambda: self._log("[Tab6] Merge worker started"))
        worker.signals.progress.connect(self._on_progress)
        worker.signals.finished.connect(self._on_merge_finished)
        worker.signals.error.connect(self._on_merge_error)

        self.thread_pool.start(worker)

    def _on_merge_finished(self, result):
        self.merge_worker = None
        self.btn_merge.setEnabled(True)

        output_pdf = Path(result)
        self._state_set("tab8_final_pdf", output_pdf)

        total_pages = self._get_checked_pages_total()

        self._set_progress_value(100, "Completed")
        self._log(f"[Tab6] Final merged PDF: {output_pdf}")
        self._log(f"[Tab6] Total merged pages: {total_pages}")

        QMessageBox.information(
            self,
            "PDF Merge Complete",
            f"Final PDF created:\n{output_pdf}\n\n"
            f"Total merged pages: {total_pages}",
        )

    def _on_merge_error(self, error_text: str):
        self.merge_worker = None
        self.btn_merge.setEnabled(True)
        self._set_progress_idle()
        self.progress_status_label.setText("Failed")

        self._log("[Tab6] MERGE ERROR")
        self._log(error_text)

        QMessageBox.critical(self, "Merge Failed", error_text)



Tab8Widget = PdfWidget
PdfMergeWidget = PdfWidget
PptxToPdfWidget = PdfWidget
PPTXToPDFWidget = PdfWidget

# =========================================================
# DIRECT TEST RUNNER
# =========================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = PdfWidget()
    window.resize(1350, 750)
    window.show()

    sys.exit(app.exec())