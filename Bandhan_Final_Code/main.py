from __future__ import annotations

import importlib
import inspect
import json
import shutil
import sys
from pathlib import Path
from typing import Any

from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

# =========================================================
# EXCEL FILES GENERATOR TAB
# =========================================================
# excel_files_generator.py should be inside the tabs folder:
# tabs/excel_files_generator.py
from tabs.excel_files_generator import ExcelFilesGeneratorTab


# =========================================================
# APP PATH HELPERS
# =========================================================
def get_app_run_folder() -> Path:
    """Return the folder where the application is running."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_output_root() -> Path:
    folder = get_app_run_folder() / "Bandhan_Output"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


# =========================================================
# TAB SCOPED STATE
# =========================================================
class TabScopedState:
    """A small state proxy passed to each loaded tab.

    Why this exists:
    Some older tab widgets create folders by doing:
        state.output_root / "02_asset_allocation_excel"

    If every tab receives the global output root, those old section folders are
    created directly inside Bandhan_Output. This proxy gives each tab its own
    output_root, for example:
        Tab 2 receives Bandhan_Output/02_fund_pages_automation
        Tab 4 receives Bandhan_Output/04_snapshot_automation

    Therefore even older hardcoded section folders become nested correctly.
    """

    def __init__(self, parent: AppState, scoped_output_root: Path):  # type: ignore[name-defined]
        object.__setattr__(self, "_parent", parent)
        object.__setattr__(self, "output_root", scoped_output_root)
        scoped_output_root.mkdir(parents=True, exist_ok=True)

        # Copy parent attributes so older tab code can read them.
        for key, value in parent.__dict__.items():
            if key != "output_root":
                object.__setattr__(self, key, value)

    def __getattr__(self, name: str) -> Any:
        return getattr(object.__getattribute__(self, "_parent"), name)

    def __setattr__(self, name: str, value: Any) -> None:
        # Keep mutations local to the tab proxy. This avoids one tab changing the
        # global output_root for all other tabs.
        object.__setattr__(self, name, value)


# =========================================================
# APP STATE
# =========================================================
class AppState:
    """Shared root state.

    Clean folder rule:
    - Only main tab folders are created directly inside Bandhan_Output.
    - Each tab section folder is created inside its tab folder.
    - Known old root-level section folders are deleted automatically.
    """

    OLD_ROOT_FOLDER_NAMES = {
        "01_performance_table_note",
        "01_pptx_pdf_merger",
        "01_pptx_merge",
        "02_age_based_notes",
        "02_asset_allocation_excel",
        "02_fund_page_note",
        "03_exitload_updated",
        "03_fund_manager_count",
        "03_fund_manager_count",
        "03_hybrid_asset_snapshot_updated",
        "03_manual_fund_page_note",
        "04_exitload_updated",
        "04_sector_allocation_updated",
        "04_sip_performance_check",
        "05_gold_silver_etf_updated",
        "05_marketcap_updated",
        "05_sector_allocation_updated",
        "05_Snapshot",
        "06_gold_silver_etf_updated",
        "06_search_replace",
        "07_ppt_page_number",
        "08_pdf",
        "age_based_outputs",
    }

    MAIN_ROOT_FOLDER_NAMES = {
        "Excel_Files_Generator_Output",
        "01_fund_manager_performance",
        "02_fund_pages_automation",
        "03_performance_check",
        "04_snapshot_automation",
        "05_search_replace",
        "06_ppt_page_number",
        "07_pdf",
    }

    def __init__(self) -> None:
        self.output_root: Path | None = default_output_root()

        # Excel Files Generator root folder.
        self.excel_files_generator_root_folder: Path | None = None
        self.excel_files_generator_output_folder: Path | None = None

        # Main root folders.
        self.tab1_root_folder: Path | None = None
        self.tab2_root_folder: Path | None = None
        self.tab3_root_folder: Path | None = None
        self.tab4_root_folder: Path | None = None
        self.tab5_root_folder: Path | None = None
        self.tab6_root_folder: Path | None = None
        self.tab7_root_folder: Path | None = None

        # Compatibility names from older versions.
        self.tab1_output_folder: Path | None = None
        self.tab2_output_folder: Path | None = None
        self.tab3_output_folder: Path | None = None
        self.tab4_output_folder: Path | None = None
        self.tab5_updated_folder: Path | None = None
        self.tab6_output_folder: Path | None = None
        self.tab7_output_folder: Path | None = None
        self.tab8_pdf_folder: Path | None = None
        self.tab8_final_pdf: Path | None = None
        self.tab6_pdf_folder: Path | None = None
        self.tab6_final_pdf: Path | None = None
        self.tab3_split_folder: Path | None = None
        self.tab4_report_folder: Path | None = None
        self.tab5_report_folder: Path | None = None

        # Tab 1 sections.
        self.tab1_section1_performance_table_note_folder: Path | None = None
        self.tab1_section2_fund_manager_count_folder: Path | None = None

        # Tab 2 sections.
        self.tab2_section1_pptx_pdf_merger_folder: Path | None = None
        self.tab2_section1_pptx_merge_folder: Path | None = None
        self.tab2_section2_age_based_notes_folder: Path | None = None
        self.tab2_section3_manual_fund_page_note_folder: Path | None = None
        self.tab2_section4_exitload_folder: Path | None = None
        self.tab2_section5_sector_allocation_folder: Path | None = None
        self.tab2_section6_gold_silver_etf_folder: Path | None = None

        # Tab 3 sections.
        self.tab3_section1_sip_performance_check_folder: Path | None = None
        self.tab3_section2_performance_scheme_check_folder: Path | None = None

        # Tab 4 sections.
        self.tab4_section1_marketcap_folder: Path | None = None
        self.tab4_section2_asset_allocation_excel_folder: Path | None = None
        self.tab4_section3_hybrid_asset_snapshot_folder: Path | None = None

        # Single-section tabs.
        self.tab5_search_replace_folder: Path | None = None
        self.tab6_ppt_page_number_folder: Path | None = None
        self.tab7_pdf_folder: Path | None = None
        self.tab7_final_pdf: Path | None = None

        self.initialize_default_output_folders()

    def ensure_output_subdir(self, name: str) -> Path:
        if self.output_root is None:
            self.output_root = default_output_root()
        folder = self.output_root / name
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    @staticmethod
    def ensure_child(parent: Path, child_name: str) -> Path:
        folder = parent / child_name
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def cleanup_old_root_output_folders(self) -> None:
        if self.output_root is None:
            return
        self.output_root.mkdir(parents=True, exist_ok=True)

        for old_name in sorted(self.OLD_ROOT_FOLDER_NAMES):
            old_path = self.output_root / old_name
            if old_path.exists() and old_path.is_dir() and old_name not in self.MAIN_ROOT_FOLDER_NAMES:
                shutil.rmtree(old_path, ignore_errors=True)

    def cleanup_tab2_duplicate_child_folders(self) -> None:
        """Remove old duplicated Tab 2 child folders created by earlier builds."""
        if self.tab2_root_folder is None:
            return
        old_child_names = {
            "01_pptx_merge",
            "02_fund_page_note",
            "03_exitload_updated",
            "05_gold_silver_etf_updated",
            "age_based_outputs",
        }
        for child_name in sorted(old_child_names):
            child_path = self.tab2_root_folder / child_name
            if child_path.exists() and child_path.is_dir():
                shutil.rmtree(child_path, ignore_errors=True)

    def cleanup_tab4_duplicate_child_folders(self) -> None:
        """Remove old duplicated Tab 4 child folders created by earlier builds."""
        if self.tab4_root_folder is None:
            return
        old_child_names = {
            "05_marketcap_updated",
            "05_Snapshot",
        }
        for child_name in sorted(old_child_names):
            child_path = self.tab4_root_folder / child_name
            if child_path.exists() and child_path.is_dir():
                shutil.rmtree(child_path, ignore_errors=True)

    def initialize_default_output_folders(self) -> None:
        if self.output_root is None:
            self.output_root = default_output_root()
        self.output_root.mkdir(parents=True, exist_ok=True)

        # Delete old root-level folders first.
        self.cleanup_old_root_output_folders()

        # Excel Files Generator main folder.
        self.excel_files_generator_root_folder = self.ensure_output_subdir("Excel_Files_Generator_Output")
        self.excel_files_generator_output_folder = self.excel_files_generator_root_folder

        # Main tab folders.
        self.tab1_root_folder = self.ensure_output_subdir("01_fund_manager_performance")
        self.tab2_root_folder = self.ensure_output_subdir("02_fund_pages_automation")
        self.tab3_root_folder = self.ensure_output_subdir("03_performance_check")
        self.tab4_root_folder = self.ensure_output_subdir("04_snapshot_automation")
        self.tab5_root_folder = self.ensure_output_subdir("05_search_replace")
        self.tab6_root_folder = self.ensure_output_subdir("06_ppt_page_number")
        self.tab7_root_folder = self.ensure_output_subdir("07_pdf")

        # Remove duplicated old Tab 2 child folders before creating the final structure.
        self.cleanup_tab2_duplicate_child_folders()

        # Remove duplicated old Tab 4 child folders before creating the final structure.
        self.cleanup_tab4_duplicate_child_folders()

        # Tab 1 section folders.
        self.tab1_section1_performance_table_note_folder = self.ensure_child(
            self.tab1_root_folder,
            "01_performance_table_note",
        )
        self.tab1_section2_fund_manager_count_folder = self.ensure_child(
            self.tab1_root_folder,
            "02_fund_manager_count",
        )

        # Tab 2 section folders.
        self.tab2_section1_pptx_pdf_merger_folder = self.ensure_child(
            self.tab2_root_folder,
            "01_pptx_pdf_merger",
        )
        self.tab2_section2_age_based_notes_folder = self.ensure_child(
            self.tab2_root_folder,
            "02_age_based_notes",
        )
        self.tab2_section3_manual_fund_page_note_folder = self.ensure_child(
            self.tab2_root_folder,
            "03_manual_fund_page_note",
        )
        self.tab2_section4_exitload_folder = self.ensure_child(
            self.tab2_root_folder,
            "04_exitload_updated",
        )
        self.tab2_section5_sector_allocation_folder = self.ensure_child(
            self.tab2_root_folder,
            "05_sector_allocation_updated",
        )
        self.tab2_section6_gold_silver_etf_folder = self.ensure_child(
            self.tab2_root_folder,
            "06_gold_silver_etf_updated",
        )

        # Tab 3 section folders.
        self.tab3_section1_sip_performance_check_folder = self.ensure_child(
            self.tab3_root_folder,
            "01_sip_performance_check",
        )
        self.tab3_section2_performance_scheme_check_folder = self.ensure_child(
            self.tab3_root_folder,
            "02_performance_scheme_check",
        )

        # Tab 4 section folders.
        self.tab4_section1_marketcap_folder = self.ensure_child(
            self.tab4_root_folder,
            "01_marketcap_updated",
        )
        self.tab4_section2_asset_allocation_excel_folder = self.ensure_child(
            self.tab4_root_folder,
            "02_asset_allocation_excel",
        )
        self.tab4_section3_hybrid_asset_snapshot_folder = self.ensure_child(
            self.tab4_root_folder,
            "03_hybrid_asset_snapshot_updated",
        )

        # Single-section tabs.
        self.tab5_search_replace_folder = self.tab5_root_folder
        self.tab6_ppt_page_number_folder = self.tab6_root_folder
        self.tab7_pdf_folder = self.tab7_root_folder
        self.tab7_final_pdf = self.tab7_pdf_folder / "Bandhan_Final_Factsheet.pdf"

        # Compatibility mappings.
        self.tab1_output_folder = self.tab1_root_folder
        self.tab2_output_folder = self.tab2_root_folder
        self.tab3_output_folder = self.tab3_root_folder
        self.tab4_output_folder = self.tab4_root_folder

        self.tab5_updated_folder = self.tab4_section1_marketcap_folder
        self.tab5_marketcap_folder = self.tab4_section1_marketcap_folder
        self.tab5_asset_allocation_excel_folder = self.tab4_section2_asset_allocation_excel_folder
        self.tab5_hybrid_asset_snapshot_folder = self.tab4_section3_hybrid_asset_snapshot_folder

        self.tab6_output_folder = self.tab5_search_replace_folder
        self.tab7_output_folder = self.tab6_ppt_page_number_folder
        self.tab8_pdf_folder = self.tab7_pdf_folder
        self.tab8_final_pdf = self.tab7_final_pdf
        self.tab6_pdf_folder = self.tab7_pdf_folder
        self.tab6_final_pdf = self.tab7_final_pdf

        self.tab3_split_folder = self.tab1_section2_fund_manager_count_folder
        self.tab4_report_folder = self.tab3_root_folder
        self.tab5_report_folder = self.tab4_root_folder

        self.tab2_pptx_pdf_merger_folder = self.tab2_section1_pptx_pdf_merger_folder
        self.tab2_pptx_merge_folder = self.tab2_section1_pptx_pdf_merger_folder
        self.tab2_section1_pptx_merge_folder = self.tab2_section1_pptx_pdf_merger_folder
        self.tab2_age_based_folder = self.tab2_section2_age_based_notes_folder
        self.tab2_manual_note_folder = self.tab2_section3_manual_fund_page_note_folder
        self.tab2_exitload_folder = self.tab2_section4_exitload_folder
        self.tab2_sector_allocation_folder = self.tab2_section5_sector_allocation_folder
        self.tab2_gold_silver_etf_folder = self.tab2_section6_gold_silver_etf_folder
        # Legacy section aliases intentionally point to final Section 4/6 folders.
        self.tab2_section3_exitload_folder = self.tab2_section4_exitload_folder
        self.tab2_section5_gold_silver_etf_folder = self.tab2_section6_gold_silver_etf_folder

    def scoped_state_for_tab(self, tab_number: int) -> TabScopedState:
        root_map = {
            0: self.excel_files_generator_root_folder,
            1: self.tab1_root_folder,
            2: self.tab2_root_folder,
            3: self.tab3_root_folder,
            4: self.tab4_root_folder,
            5: self.tab5_root_folder,
            6: self.tab6_root_folder,
            7: self.tab7_root_folder,
        }
        scoped_root = root_map.get(tab_number) or self.output_root or default_output_root()
        return TabScopedState(self, scoped_root)

    def reset_output_root(self, folder: Path) -> None:
        self.output_root = Path(folder)
        self.output_root.mkdir(parents=True, exist_ok=True)
        self.initialize_default_output_folders()

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in self.__dict__.items():
            result[key] = str(value) if isinstance(value, Path) else value
        return result

    def load_dict(self, data: dict[str, Any]) -> None:
        for key, value in data.items():
            if value is None:
                setattr(self, key, None)
            elif (
                key == "output_root"
                or key.endswith("_folder")
                or key.endswith("_pdf")
                or key.endswith("_file")
                or key.endswith("_path")
            ):
                setattr(self, key, Path(value))
            else:
                setattr(self, key, value)

        if self.output_root is None:
            self.output_root = default_output_root()
        self.output_root.mkdir(parents=True, exist_ok=True)
        self.initialize_default_output_folders()


# =========================================================
# TAB LOADER
# =========================================================
def _instantiate_widget(widget_class: type[QWidget], state: Any) -> QWidget:
    try:
        return widget_class(state)
    except TypeError:
        return widget_class()


def load_tab_widget(module_name: str, state: Any) -> QWidget:
    module = importlib.import_module(module_name)

    preferred_names = [
        "Tab1Widget", "Tab2Widget", "Tab3Widget", "Tab4Widget",
        "Tab5Widget", "Tab6Widget", "Tab7Widget", "Tab8Widget",
        "CombinedPerformanceFundManagerWidget",
        "CombinedFundPageNotesExitLoadSectorGoldSilverWidget",
        "CombinedFundPageNotesExitLoadSectorWidget",
        "CombinedChecksWidget", "CombinedChecksWindow",
        "CombinedMarketAssetHybridWidget",
        "PerformanceTableNoteWindow", "FundPageNoteWindow", "FundManagerPPTChecker",
        "PerformanceTableNoteWidget", "FundManagerCountWidget",
        "GoldSilverETFSectionWidget", "ExitLoadSectionWidget", "SectorSectionWidget",
        "MarketCapWidget", "SnapshotMarketCapUpdater", "SnapshotMarketCapWidget",
        "SearchReplaceWidget", "SearchAndReplaceWidget",
        "UpdateRefsWidget", "UpdateRefsStandaloneWidget",
        "PdfMergeWidget", "PptxToPdfWidget", "PPTXToPDFWidget", "PdfWidget",
        "MainWidget", "MainWindow",
    ]

    for name in preferred_names:
        obj = getattr(module, name, None)
        if obj is None:
            continue
        try:
            if inspect.isclass(obj) and issubclass(obj, QWidget):
                return _instantiate_widget(obj, state)
        except TypeError:
            continue

    skip_class_names = {
        "QWidget", "QMainWindow", "QTableWidget", "QStyledItemDelegate",
        "PptTable", "RowBorderDelegate", "WorkerSignals", "FunctionWorker",
        "Table", "Delegate",
    }

    widget_classes: list[type[QWidget]] = []
    for name, obj in inspect.getmembers(module, inspect.isclass):
        if obj is QWidget:
            continue
        if obj.__module__ != module.__name__:
            continue
        if name in skip_class_names or name.endswith("Table") or name.endswith("Delegate"):
            continue
        try:
            if issubclass(obj, QWidget):
                widget_classes.append(obj)
        except TypeError:
            continue

    if not widget_classes:
        raise RuntimeError(f"No valid QWidget tab class found in module:\n{module_name}")
    return _instantiate_widget(widget_classes[0], state)


# =========================================================
# MAIN WINDOW
# =========================================================
class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.state = AppState()
        self.setWindowTitle("Bandhan Factsheet Automation")
        self.resize(1365, 760)
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(6, 6, 6, 6)
        root_layout.setSpacing(6)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        self.btn_choose_output_root = QPushButton("Choose Output Root")
        self.btn_save_state = QPushButton("Save State")
        self.btn_load_state = QPushButton("Load State")
        self.output_root_label = QLabel()
        self._refresh_output_root_label()

        top_row.addWidget(self.btn_choose_output_root)
        top_row.addWidget(self.btn_save_state)
        top_row.addWidget(self.btn_load_state)
        top_row.addWidget(self.output_root_label)
        top_row.addStretch(1)
        root_layout.addLayout(top_row)

        self.tabs = QTabWidget()
        root_layout.addWidget(self.tabs, stretch=1)

        self._load_all_tabs()

        self.btn_choose_output_root.clicked.connect(self._choose_output_root)
        self.btn_save_state.clicked.connect(self._save_state)
        self.btn_load_state.clicked.connect(self._load_state)

    def _load_all_tabs(self) -> None:
        self.tabs.clear()

        # Excel Files Generator is added as the first tab without removing existing tabs.
        self._add_excel_files_generator_tab()

        self._add_tab("tabs.tab1_performance", "Tab 1 - Fund Manager Performance", 1)
        self._add_tab("tabs.tab2_fundpage", "Tab 2 - Fund Pages Automation", 2)
        self._add_tab("tabs.tab3_performancecheck", "Tab 3 - Performance Check", 3)
        self._add_tab("tabs.tab4_snapshot", "Tab 4 - Snapshot Automation", 4)
        self._add_tab("tabs.tab5_searchandreplace", "Tab 5 - Search / Replace", 5)
        self._add_tab("tabs.tab6_pptpagenumber", "Tab 6 - PPT Page Number Adder", 6)
        self._add_tab("tabs.tab7_pdf", "Tab 7 - PPTX to PDF / Merge", 7)

        # Some older widgets create old folders during their __init__.
        # Clean again after all tabs are loaded.
        self.state.cleanup_old_root_output_folders()

    def _refresh_output_root_label(self) -> None:
        self.output_root_label.setText(f"Output Root: {self.state.output_root}")

    def _add_excel_files_generator_tab(self) -> None:
        try:
            scoped_state = self.state.scoped_state_for_tab(0)
            widget = ExcelFilesGeneratorTab(scoped_state)
            self.tabs.addTab(widget, "Excel Files Generator")
        except Exception as exc:
            self.tabs.addTab(
                self._make_error_tab(
                    "Excel Files Generator",
                    "tabs.excel_files_generator",
                    exc,
                ),
                "Excel Files Generator",
            )

    def _add_tab(self, module_name: str, title: str, tab_number: int) -> None:
        try:
            scoped_state = self.state.scoped_state_for_tab(tab_number)
            widget = load_tab_widget(module_name, scoped_state)
            self.tabs.addTab(widget, title)
        except Exception as exc:
            self.tabs.addTab(self._make_error_tab(title, module_name, exc), title)

    def _make_error_tab(self, title: str, module_name: str, exc: Exception) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        label = QLabel(
            f"Could not load tab:\n\n{title}\n\nModule:\n{module_name}\n\nError:\n{exc}"
        )
        label.setWordWrap(True)
        layout.addWidget(label)
        layout.addStretch(1)
        return widget

    def _choose_output_root(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Choose Output Root Folder",
            str(self.state.output_root or default_output_root()),
        )
        if not folder:
            return
        self.state.reset_output_root(Path(folder))
        self._refresh_output_root_label()
        self._load_all_tabs()
        QMessageBox.information(
            self,
            "Output Root Selected",
            f"Output root folder selected:\n\n{self.state.output_root}",
        )

    def _save_state(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save State",
            str((self.state.output_root or default_output_root()) / "app_state.json"),
            "JSON Files (*.json)",
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as file:
                json.dump(self.state.to_dict(), file, indent=2)
            QMessageBox.information(self, "State Saved", f"State saved successfully:\n{file_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Save Failed", str(exc))

    def _load_state(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Load State",
            str(self.state.output_root or default_output_root()),
            "JSON Files (*.json)",
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                data = json.load(file)
            self.state.load_dict(data)
            self._refresh_output_root_label()
            self._load_all_tabs()
            QMessageBox.information(self, "State Loaded", f"State loaded successfully:\n{file_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Load Failed", str(exc))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())