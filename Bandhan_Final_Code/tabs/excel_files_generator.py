from __future__ import annotations

import os
import shutil
import sys
from copy import copy
from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from openpyxl import load_workbook


# =========================================================
# PATH HELPERS
# =========================================================

def get_app_run_folder() -> Path:
    """
    Returns exe folder when frozen by PyInstaller,
    otherwise returns python file folder.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_excel_generator_output_folder(state: Any = None) -> Path:
    """
    Creates/returns Excel Files Generator output folder inside Bandhan_Output.

    When opened from main.py:
        state.output_root is scoped to Bandhan_Output/Excel_Files_Generator_Output

    When opened standalone:
        fallback is <project folder>/Bandhan_Output/Excel_Files_Generator_Output

    This keeps Excel Generator output folder consistent with all other tabs.
    """
    if state is not None and hasattr(state, "output_root") and getattr(state, "output_root") is not None:
        folder = Path(getattr(state, "output_root"))
    else:
        app_folder = get_app_run_folder()

        # If this file is inside tabs/, project folder is parent of tabs/.
        if app_folder.name.lower() == "tabs":
            app_folder = app_folder.parent

        folder = app_folder / "Bandhan_Output" / "Excel_Files_Generator_Output"

    folder.mkdir(parents=True, exist_ok=True)
    return folder


def make_output_copy(input_file: str, output_folder: str, suffix: str = "_updated") -> str:
    """
    Creates a copy of target Excel file into output folder.
    If file exists, creates incremented filename.
    """
    src = Path(input_file)
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    base_name = src.stem
    ext = src.suffix

    output_path = out_dir / f"{base_name}{suffix}{ext}"

    counter = 1
    while output_path.exists():
        output_path = out_dir / f"{base_name}{suffix}_{counter}{ext}"
        counter += 1

    shutil.copy2(src, output_path)
    return str(output_path)


# =========================================================
# COMMON BASE
# =========================================================

class ExcelToolBase(QWidget):
    def __init__(self, state: Any = None, output_folder_getter=None):
        super().__init__()
        self.state = state
        self.output_folder_getter = output_folder_getter

    def log(self, message: str):
        if hasattr(self, "log_box"):
            self.log_box.append(message)

    def is_blank(self, value) -> bool:
        if value is None:
            return True
        if str(value).strip() == "":
            return True
        return False

    def find_headers(self, worksheet, header_row: int = 1) -> dict:
        headers = {}
        for cell in worksheet[header_row]:
            if cell.value is not None:
                headers[str(cell.value).strip()] = cell.column
        return headers

    def copy_cell_style(self, source_cell, target_cell):
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)

    def clear_target_data(self, worksheet, start_row: int = 2):
        max_row = worksheet.max_row
        max_col = worksheet.max_column

        for row in range(start_row, max_row + 1):
            for col in range(1, max_col + 1):
                worksheet.cell(row=row, column=col).value = None

    def get_output_folder(self) -> str:
        if self.output_folder_getter:
            folder = self.output_folder_getter()
            if folder:
                Path(folder).mkdir(parents=True, exist_ok=True)
                return folder

        folder = default_excel_generator_output_folder(self.state)
        return str(folder)

    def file_picker_row(self, label_text, line_edit, browse_function, label_width=230):
        row = QHBoxLayout()

        label = QLabel(label_text)
        label.setFixedWidth(label_width)

        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(browse_function)

        row.addWidget(label)
        row.addWidget(line_edit)
        row.addWidget(browse_btn)

        return row

    def show_error_notification(self, error_message: str):
        QMessageBox.critical(
            self,
            "Error",
            f"Process failed due to the following error:\n\n{error_message}"
        )


# =========================================================
# SECTION 1 - FUND MANAGER DETAILS
# =========================================================

class FundManagerSection(ExcelToolBase):
    def __init__(self, state: Any = None, output_folder_getter=None):
        super().__init__(state, output_folder_getter)

        self.uploader_path = QLineEdit()
        self.target_path = QLineEdit()

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        title = QLabel("Section 1 - Fund Manager Details")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title)

        main_layout.addLayout(
            self.file_picker_row(
                "Uploader Excel:",
                self.uploader_path,
                self.browse_uploader,
                label_width=230
            )
        )

        main_layout.addLayout(
            self.file_picker_row(
                "FundManagerdetail Excel:",
                self.target_path,
                self.browse_target,
                label_width=230
            )
        )

        note = QLabel(
            "Output: A copied updated FundManagerdetail file will be created in the selected output folder."
        )
        note.setStyleSheet("color: #B00020; font-weight: bold;")
        main_layout.addWidget(note)

        run_btn = QPushButton("Run Fund Manager Details")
        run_btn.setStyleSheet(
            """
            QPushButton {
                font-size: 14px;
                padding: 8px;
                font-weight: bold;
                background-color: #0078D4;
                color: white;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #005A9E;
            }
            """
        )
        run_btn.clicked.connect(self.run_process)
        main_layout.addWidget(run_btn)

        main_layout.addWidget(QLabel("Process Log / Missing Data Details:"))
        main_layout.addWidget(self.log_box)

        self.setLayout(main_layout)

    def browse_uploader(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Uploader Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm)"
        )

        if file_path:
            self.uploader_path.setText(file_path)

    def browse_target(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select FundManagerdetail Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm)"
        )

        if file_path:
            self.target_path.setText(file_path)

    def show_missing_data_notification(self, missing_data):
        if not missing_data:
            QMessageBox.information(
                self,
                "Missing Data Check",
                "No missing data found in selected columns."
            )
            return

        total_missing = len(missing_data)
        preview_lines = []

        for item in missing_data[:20]:
            preview_lines.append(
                f"Fund: {item['fund_name']} | Missing Column: {item['column_name']}"
            )

        preview_text = "\n".join(preview_lines)

        if total_missing > 20:
            preview_text += (
                f"\n\n...and {total_missing - 20} more missing values. "
                f"Please check the log box for full details."
            )

        QMessageBox.warning(
            self,
            "Missing Data Found",
            f"Missing data found in {total_missing} place(s).\n\n{preview_text}"
        )

    def show_success_notification(self, copied_count, output_file, missing_count):
        QMessageBox.information(
            self,
            "Success",
            f"Fund Manager Details completed successfully.\n\n"
            f"Rows copied: {copied_count}\n"
            f"Missing data count: {missing_count}\n\n"
            f"Output file:\n{output_file}"
        )

    def run_process(self):
        uploader_file = self.uploader_path.text().strip()
        target_file = self.target_path.text().strip()

        if not uploader_file:
            QMessageBox.warning(self, "Missing File", "Please select Uploader Excel file.")
            return

        if not target_file:
            QMessageBox.warning(self, "Missing File", "Please select FundManagerdetail Excel file.")
            return

        if not os.path.exists(uploader_file):
            QMessageBox.warning(self, "File Not Found", "Uploader Excel file does not exist.")
            return

        if not os.path.exists(target_file):
            QMessageBox.warning(self, "File Not Found", "FundManagerdetail Excel file does not exist.")
            return

        try:
            self.log_box.clear()
            self.log("Process started...")
            self.log("Please make sure both Excel files are closed before running.")

            output_folder = self.get_output_folder()
            output_file = make_output_copy(
                input_file=target_file,
                output_folder=output_folder,
                suffix="_updated"
            )

            self.log(f"Output folder: {output_folder}")
            self.log(f"Created output copy: {output_file}")

            uploader_sheet_name = "Uploader"
            target_sheet_name = "FundManagerDetails"

            self.log(f"Loading uploader file: {uploader_file}")
            uploader_wb = load_workbook(uploader_file, data_only=False)

            self.log(f"Loading copied FundManagerdetail file: {output_file}")
            target_wb = load_workbook(output_file)

            if uploader_sheet_name not in uploader_wb.sheetnames:
                raise Exception(f"Sheet '{uploader_sheet_name}' not found in uploader file.")

            if target_sheet_name not in target_wb.sheetnames:
                raise Exception(f"Sheet '{target_sheet_name}' not found in FundManagerdetail file.")

            uploader_ws = uploader_wb[uploader_sheet_name]
            target_ws = target_wb[target_sheet_name]

            uploader_headers = self.find_headers(uploader_ws, header_row=1)
            target_headers = self.find_headers(target_ws, header_row=1)

            column_mapping = {
                "Scheme Name": "Factsheet Display Name",
                "Fund Manager": "Fund Manager",
                "Fund Manager Display": "Fund Manager Display",
                "Managing Since": "Managing Since",
                "Launch Date Regular(DDMMYYY)": "Launch Date Regular(DDMMYYY)",
                "Launch Date Direct(DDMMYYY)": "Launch Date Direct(DDMMYYY)"
            }

            alternate_source_headers = {
                "Launch Date Regular(DDMMYYY)": [
                    "Launch Date Regular(DDMMYYY)",
                    "Launch Date Regular(DDMMYYYY)"
                ],
                "Launch Date Direct(DDMMYYY)": [
                    "Launch Date Direct(DDMMYYY)",
                    "Launch Date Direct(DDMMYYYY)"
                ]
            }

            self.log("Checking target columns...")

            for target_col in column_mapping.keys():
                if target_col not in target_headers:
                    raise Exception(f"Target column not found: {target_col}")

            self.log("Checking source columns...")

            resolved_mapping = {}

            for target_col, source_col in column_mapping.items():
                possible_source_cols = alternate_source_headers.get(
                    source_col,
                    [source_col]
                )

                found_source_col = None

                for possible_col in possible_source_cols:
                    if possible_col in uploader_headers:
                        found_source_col = possible_col
                        break

                if not found_source_col:
                    raise Exception(
                        f"Source column not found for target column '{target_col}'. "
                        f"Tried source columns: {possible_source_cols}"
                    )

                resolved_mapping[target_col] = found_source_col

            self.log("Column mapping:")
            for target_col, source_col in resolved_mapping.items():
                self.log(f"  {target_col} <- {source_col}")

            style_template = {}

            for target_col_name in column_mapping.keys():
                target_col_num = target_headers[target_col_name]
                style_template[target_col_name] = copy(
                    target_ws.cell(row=2, column=target_col_num)._style
                )

            self.log("Clearing old data from FundManagerDetails sheet...")
            self.clear_target_data(target_ws, start_row=2)

            for target_col_name in column_mapping.keys():
                target_col_num = target_headers[target_col_name]
                target_ws.cell(row=1, column=target_col_num).value = target_col_name

            missing_data = []
            copied_count = 0
            output_row = 2

            uploader_max_row = uploader_ws.max_row

            scheme_source_header = resolved_mapping["Scheme Name"]
            scheme_source_col = uploader_headers[scheme_source_header]

            self.log("Copying data and checking missing values...")

            for source_row in range(2, uploader_max_row + 1):
                scheme_value = uploader_ws.cell(
                    row=source_row,
                    column=scheme_source_col
                ).value

                if self.is_blank(scheme_value):
                    continue

                fund_name = str(scheme_value).strip()

                for target_col_name, source_col_name in resolved_mapping.items():
                    source_col_num = uploader_headers[source_col_name]
                    target_col_num = target_headers[target_col_name]

                    source_cell = uploader_ws.cell(
                        row=source_row,
                        column=source_col_num
                    )

                    target_cell = target_ws.cell(
                        row=output_row,
                        column=target_col_num
                    )

                    value = source_cell.value
                    target_cell.value = value

                    if target_col_name in style_template:
                        target_cell._style = copy(style_template[target_col_name])

                    target_cell.number_format = source_cell.number_format

                    if self.is_blank(value):
                        missing_data.append(
                            {
                                "fund_name": fund_name,
                                "column_name": target_col_name,
                                "excel_row": source_row
                            }
                        )

                copied_count += 1
                output_row += 1

            self.log(f"Rows copied: {copied_count}")

            if missing_data:
                self.log("")
                self.log("Missing Data Found:")
                self.log("-------------------")

                for item in missing_data:
                    self.log(
                        f"Fund: {item['fund_name']} | "
                        f"Missing Column: {item['column_name']} | "
                        f"Uploader Row: {item['excel_row']}"
                    )
            else:
                self.log("")
                self.log("No missing data found in selected columns.")

            self.log("")
            self.log("Saving output file...")
            target_wb.save(output_file)

            uploader_wb.close()
            target_wb.close()

            self.log("FundManagerdetail output file created successfully.")

            self.show_missing_data_notification(missing_data)

            self.show_success_notification(
                copied_count=copied_count,
                output_file=output_file,
                missing_count=len(missing_data)
            )

        except PermissionError:
            error_message = (
                "Permission denied while saving the output file.\n\n"
                "Please close the Excel file if it is already open, then run again."
            )
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)

        except Exception as e:
            error_message = str(e)
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)


# =========================================================
# SECTION 2 - CONSOLIDATE PORTFOLIO
# =========================================================

class ConsolidatePortfolioSection(ExcelToolBase):
    def __init__(self, state: Any = None, output_folder_getter=None):
        super().__init__(state, output_folder_getter)

        self.client_file_path = QLineEdit()
        self.my_file_path = QLineEdit()

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        title = QLabel("Section 2 - Consolidate Portfolio")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title)

        main_layout.addLayout(
            self.file_picker_row(
                "Client Month File:",
                self.client_file_path,
                self.browse_client_file,
                label_width=230
            )
        )

        main_layout.addLayout(
            self.file_picker_row(
                "Your ConsolidatePortfolio File:",
                self.my_file_path,
                self.browse_my_file,
                label_width=230
            )
        )

        note = QLabel(
            "Output: A copied updated ConsolidatePortfolio file will be created in the selected output folder."
        )
        note.setStyleSheet("color: #B00020; font-weight: bold;")
        main_layout.addWidget(note)

        run_btn = QPushButton("Run Consolidate Portfolio")
        run_btn.setStyleSheet(
            """
            QPushButton {
                font-size: 14px;
                padding: 8px;
                font-weight: bold;
                background-color: #0078D4;
                color: white;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #005A9E;
            }
            """
        )
        run_btn.clicked.connect(self.run_process)
        main_layout.addWidget(run_btn)

        main_layout.addWidget(QLabel("Process Log / Missing Data Details:"))
        main_layout.addWidget(self.log_box)

        self.setLayout(main_layout)

    def browse_client_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Client ConsolidatePortfolio File",
            "",
            "Excel Files (*.xlsx *.xlsm)"
        )

        if file_path:
            self.client_file_path.setText(file_path)

    def browse_my_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Your ConsolidatePortfolio File",
            "",
            "Excel Files (*.xlsx *.xlsm)"
        )

        if file_path:
            self.my_file_path.setText(file_path)

    def get_source_sheet(self, workbook):
        preferred_sheets = [
            "Portfolio",
            "ConsolidatePortfolio",
            "Consolidate Portfolio",
            "Sheet1"
        ]

        for sheet_name in preferred_sheets:
            if sheet_name in workbook.sheetnames:
                return workbook[sheet_name]

        return workbook.active

    def get_target_sheet(self, workbook):
        preferred_sheets = [
            "Sheet1",
            "ConsolidatePortfolio",
            "Consolidate Portfolio",
            "Portfolio"
        ]

        for sheet_name in preferred_sheets:
            if sheet_name in workbook.sheetnames:
                return workbook[sheet_name]

        return workbook.active

    def show_missing_data_notification(self, missing_data):
        if not missing_data:
            QMessageBox.information(
                self,
                "Missing Data Check",
                "No missing data found in selected columns."
            )
            return

        total_missing = len(missing_data)
        preview_lines = []

        for item in missing_data[:20]:
            preview_lines.append(
                f"Fund: {item['fund_name']} | "
                f"Security: {item['security_name']} | "
                f"Missing Column: {item['column_name']}"
            )

        preview_text = "\n".join(preview_lines)

        if total_missing > 20:
            preview_text += (
                f"\n\n...and {total_missing - 20} more missing values. "
                f"Please check the log box for full details."
            )

        QMessageBox.warning(
            self,
            "Missing Data Found",
            f"Missing data found in {total_missing} place(s).\n\n{preview_text}"
        )

    def show_success_notification(self, copied_count, output_file, missing_count):
        QMessageBox.information(
            self,
            "Success",
            f"Consolidate Portfolio completed successfully.\n\n"
            f"Rows copied: {copied_count}\n"
            f"Missing data count: {missing_count}\n\n"
            f"Output file:\n{output_file}"
        )

    def run_process(self):
        client_file = self.client_file_path.text().strip()
        my_file = self.my_file_path.text().strip()

        if not client_file:
            QMessageBox.warning(
                self,
                "Missing File",
                "Please select the client month ConsolidatePortfolio file."
            )
            return

        if not my_file:
            QMessageBox.warning(
                self,
                "Missing File",
                "Please select your ConsolidatePortfolio file."
            )
            return

        if not os.path.exists(client_file):
            QMessageBox.warning(
                self,
                "File Not Found",
                "Client file does not exist."
            )
            return

        if not os.path.exists(my_file):
            QMessageBox.warning(
                self,
                "File Not Found",
                "Your ConsolidatePortfolio file does not exist."
            )
            return

        try:
            self.log_box.clear()
            self.log("Process started...")
            self.log("Please make sure both Excel files are closed before running.")

            output_folder = self.get_output_folder()
            output_file = make_output_copy(
                input_file=my_file,
                output_folder=output_folder,
                suffix="_updated"
            )

            self.log(f"Output folder: {output_folder}")
            self.log(f"Created output copy: {output_file}")

            required_columns = [
                "Fund Name",
                "% to Net Assets",
                "Security Name",
                "AMFI_Industry",
                "Instrument Type",
                "AMFI SECTOR"
            ]

            alternate_source_headers = {
                "Fund Name": [
                    "Fund Name",
                    "Scheme Name"
                ],
                "% to Net Assets": [
                    "% to Net Assets",
                    "% to Net Asset",
                    "% To Net Assets",
                    "% To Net Asset",
                    "Net Asset %",
                    "% Net Assets"
                ],
                "Security Name": [
                    "Security Name",
                    "Instrument Name",
                    "Name of Instrument",
                    "Name of Comp/Stk/MF/Issuer"
                ],
                "AMFI_Industry": [
                    "AMFI_Industry",
                    "AMFI Industry",
                    "AMFI_Industry "
                ],
                "Instrument Type": [
                    "Instrument Type",
                    "Instrument",
                    "Security Type"
                ],
                "AMFI SECTOR": [
                    "AMFI SECTOR",
                    "AMFI Sector",
                    "AMFI_Sector"
                ]
            }

            alternate_target_headers = {
                "Fund Name": [
                    "Fund Name",
                    "Scheme Name"
                ],
                "% to Net Assets": [
                    "% to Net Assets",
                    "% to Net Asset",
                    "% To Net Assets",
                    "% To Net Asset",
                    "Net Asset %",
                    "% Net Assets"
                ],
                "Security Name": [
                    "Security Name",
                    "Instrument Name"
                ],
                "AMFI_Industry": [
                    "AMFI_Industry",
                    "AMFI Industry",
                    "AMFI_Industry "
                ],
                "Instrument Type": [
                    "Instrument Type",
                    "Instrument",
                    "Security Type"
                ],
                "AMFI SECTOR": [
                    "AMFI SECTOR",
                    "AMFI Sector",
                    "AMFI_Sector"
                ]
            }

            self.log(f"Loading client file: {client_file}")
            client_wb = load_workbook(client_file, data_only=False)

            self.log(f"Loading copied target file: {output_file}")
            my_wb = load_workbook(output_file)

            source_ws = self.get_source_sheet(client_wb)
            target_ws = self.get_target_sheet(my_wb)

            self.log(f"Client source sheet used: {source_ws.title}")
            self.log(f"Target sheet used: {target_ws.title}")

            source_headers = self.find_headers(source_ws, header_row=1)
            target_headers = self.find_headers(target_ws, header_row=1)

            resolved_source_columns = {}
            resolved_target_columns = {}

            self.log("Resolving source columns...")

            for final_col in required_columns:
                possible_names = alternate_source_headers.get(final_col, [final_col])
                found_name = None

                for name in possible_names:
                    if name in source_headers:
                        found_name = name
                        break

                if not found_name:
                    raise Exception(
                        f"Source column not found for: {final_col}\n"
                        f"Tried: {possible_names}"
                    )

                resolved_source_columns[final_col] = found_name

            self.log("Resolving target columns...")

            for final_col in required_columns:
                possible_names = alternate_target_headers.get(final_col, [final_col])
                found_name = None

                for name in possible_names:
                    if name in target_headers:
                        found_name = name
                        break

                if not found_name:
                    raise Exception(
                        f"Target column not found for: {final_col}\n"
                        f"Tried: {possible_names}"
                    )

                resolved_target_columns[final_col] = found_name

            self.log("Column mapping:")
            for final_col in required_columns:
                self.log(
                    f"  {resolved_target_columns[final_col]} <- "
                    f"{resolved_source_columns[final_col]}"
                )

            self.log("Saving target row style template...")

            style_template = {}

            for final_col in required_columns:
                target_col_num = target_headers[resolved_target_columns[final_col]]
                style_template[final_col] = copy(
                    target_ws.cell(row=2, column=target_col_num)._style
                )

            self.log("Clearing old data from copied target file...")
            self.clear_target_data(target_ws, start_row=2)

            self.log("Writing headers in target file...")

            for final_col in required_columns:
                target_col_num = target_headers[resolved_target_columns[final_col]]
                target_cell = target_ws.cell(row=1, column=target_col_num)
                target_cell.value = final_col

            fund_source_col = source_headers[resolved_source_columns["Fund Name"]]
            security_source_col = source_headers[resolved_source_columns["Security Name"]]

            copied_count = 0
            output_row = 2
            missing_data = []

            self.log("Copying data from client file to copied target file...")

            for source_row in range(2, source_ws.max_row + 1):
                fund_value = source_ws.cell(row=source_row, column=fund_source_col).value
                security_value = source_ws.cell(row=source_row, column=security_source_col).value

                if self.is_blank(fund_value) and self.is_blank(security_value):
                    continue

                fund_name = (
                    str(fund_value).strip()
                    if not self.is_blank(fund_value)
                    else "Missing Fund Name"
                )

                security_name = (
                    str(security_value).strip()
                    if not self.is_blank(security_value)
                    else "Missing Security Name"
                )

                for final_col in required_columns:
                    source_col_name = resolved_source_columns[final_col]
                    target_col_name = resolved_target_columns[final_col]

                    source_col_num = source_headers[source_col_name]
                    target_col_num = target_headers[target_col_name]

                    source_cell = source_ws.cell(row=source_row, column=source_col_num)
                    target_cell = target_ws.cell(row=output_row, column=target_col_num)

                    target_cell.value = source_cell.value

                    if final_col in style_template:
                        target_cell._style = copy(style_template[final_col])

                    target_cell.number_format = source_cell.number_format

                    if self.is_blank(source_cell.value):
                        missing_data.append(
                            {
                                "fund_name": fund_name,
                                "security_name": security_name,
                                "column_name": final_col,
                                "source_row": source_row
                            }
                        )

                copied_count += 1
                output_row += 1

            self.log(f"Rows copied: {copied_count}")

            if missing_data:
                self.log("")
                self.log("Missing Data Found:")
                self.log("-------------------")

                for item in missing_data:
                    self.log(
                        f"Fund: {item['fund_name']} | "
                        f"Security: {item['security_name']} | "
                        f"Missing Column: {item['column_name']} | "
                        f"Client File Row: {item['source_row']}"
                    )
            else:
                self.log("")
                self.log("No missing data found in selected columns.")

            self.log("")
            self.log("Saving output file...")
            my_wb.save(output_file)

            client_wb.close()
            my_wb.close()

            self.log("ConsolidatePortfolio output file created successfully.")

            self.show_missing_data_notification(missing_data)

            self.show_success_notification(
                copied_count=copied_count,
                output_file=output_file,
                missing_count=len(missing_data)
            )

        except PermissionError:
            error_message = (
                "Permission denied while saving the output file.\n\n"
                "Please close the Excel file if it is open, then run again."
            )
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)

        except Exception as e:
            error_message = str(e)
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)


# =========================================================
# MAIN COMBINED TAB
# =========================================================

class ExcelFilesGeneratorTab(QWidget):
    """
    This widget is added as first tab in main.py.
    It contains:
    Section 1 - Fund Manager Details
    Section 2 - Consolidate Portfolio
    """

    def __init__(self, state: Any = None):
        super().__init__()

        self.state = state

        self.output_folder_path = QLineEdit()
        self.output_folder_path.setText(str(default_excel_generator_output_folder(self.state)))

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        title = QLabel("Excel Files Generator")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        main_layout.addWidget(title)

        output_row = QHBoxLayout()

        output_label = QLabel("Output Folder:")
        output_label.setFixedWidth(130)

        browse_output_btn = QPushButton("Select Output Folder")
        browse_output_btn.clicked.connect(self.browse_output_folder)

        open_output_btn = QPushButton("Open Output Folder")
        open_output_btn.clicked.connect(self.open_output_folder)

        output_row.addWidget(output_label)
        output_row.addWidget(self.output_folder_path)
        output_row.addWidget(browse_output_btn)
        output_row.addWidget(open_output_btn)

        main_layout.addLayout(output_row)

        note = QLabel(
            "Default output folder is automatically created inside Bandhan_Output as: Excel_Files_Generator_Output"
        )
        note.setStyleSheet("color: #0078D4; font-weight: bold;")
        main_layout.addWidget(note)

        self.sections_tab = QTabWidget()

        self.fund_manager_section = FundManagerSection(
            state=self.state,
            output_folder_getter=self.get_output_folder
        )

        self.consolidate_section = ConsolidatePortfolioSection(
            state=self.state,
            output_folder_getter=self.get_output_folder
        )

        self.sections_tab.addTab(self.fund_manager_section, "Section 1 - Fund Manager Details")
        self.sections_tab.addTab(self.consolidate_section, "Section 2 - Consolidate Portfolio")

        main_layout.addWidget(self.sections_tab)

        self.setLayout(main_layout)

    def get_output_folder(self) -> str:
        folder = self.output_folder_path.text().strip()

        if not folder:
            folder = str(default_excel_generator_output_folder(self.state))
            self.output_folder_path.setText(folder)

        Path(folder).mkdir(parents=True, exist_ok=True)
        return folder

    def browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            self.output_folder_path.text().strip()
        )

        if folder:
            self.output_folder_path.setText(folder)

    def open_output_folder(self):
        folder = self.get_output_folder()

        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as e:
            QMessageBox.warning(
                self,
                "Open Folder Error",
                f"Could not open output folder:\n\n{e}"
            )


# Standalone testing
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelFilesGeneratorTab()
    window.resize(1200, 760)
    window.show()
    sys.exit(app.exec())