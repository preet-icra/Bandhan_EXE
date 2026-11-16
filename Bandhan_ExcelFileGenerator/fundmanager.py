import sys
import os
from copy import copy

from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
    QTextEdit
)
from PyQt6.QtCore import Qt

from openpyxl import load_workbook


class FundManagerCopyGUI(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Fund Manager Details Copier")
        self.setMinimumWidth(850)

        self.uploader_path = QLineEdit()
        self.target_path = QLineEdit()

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        title = QLabel("Copy Fund Manager Data from Uploader Excel to FundManagerdetail Excel")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title)

        main_layout.addLayout(
            self.file_picker_row(
                "Uploader Excel:",
                self.uploader_path,
                self.browse_uploader
            )
        )

        main_layout.addLayout(
            self.file_picker_row(
                "FundManagerdetail Excel:",
                self.target_path,
                self.browse_target
            )
        )

        note = QLabel(
            "Note: Data will be copied directly into the selected FundManagerdetail Excel file."
        )
        note.setStyleSheet("color: #B00020; font-weight: bold;")
        main_layout.addWidget(note)

        run_btn = QPushButton("Run")
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

    def file_picker_row(self, label_text, line_edit, browse_function):
        row = QHBoxLayout()

        label = QLabel(label_text)
        label.setFixedWidth(180)

        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(browse_function)

        row.addWidget(label)
        row.addWidget(line_edit)
        row.addWidget(browse_btn)

        return row

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

    def log(self, message):
        self.log_box.append(message)

    def is_blank(self, value):
        if value is None:
            return True

        if str(value).strip() == "":
            return True

        return False

    def copy_cell_style(self, source_cell, target_cell):
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)

    def find_headers(self, worksheet, header_row=1):
        headers = {}

        for cell in worksheet[header_row]:
            if cell.value is not None:
                headers[str(cell.value).strip()] = cell.column

        return headers

    def clear_target_data(self, worksheet, start_row=2):
        max_row = worksheet.max_row
        max_col = worksheet.max_column

        for row in range(start_row, max_row + 1):
            for col in range(1, max_col + 1):
                worksheet.cell(row=row, column=col).value = None

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

    def show_success_notification(self, copied_count, target_file, missing_count):
        QMessageBox.information(
            self,
            "Success",
            f"Process completed successfully.\n\n"
            f"Rows copied: {copied_count}\n"
            f"Missing data count: {missing_count}\n\n"
            f"Updated file:\n{target_file}"
        )

    def show_error_notification(self, error_message):
        QMessageBox.critical(
            self,
            "Error",
            f"Process failed due to the following error:\n\n{error_message}"
        )

    def run_process(self):
        uploader_file = self.uploader_path.text().strip()
        target_file = self.target_path.text().strip()

        if not uploader_file:
            QMessageBox.warning(
                self,
                "Missing File",
                "Please select Uploader Excel file."
            )
            return

        if not target_file:
            QMessageBox.warning(
                self,
                "Missing File",
                "Please select FundManagerdetail Excel file."
            )
            return

        if not os.path.exists(uploader_file):
            QMessageBox.warning(
                self,
                "File Not Found",
                "Uploader Excel file does not exist."
            )
            return

        if not os.path.exists(target_file):
            QMessageBox.warning(
                self,
                "File Not Found",
                "FundManagerdetail Excel file does not exist."
            )
            return

        try:
            self.log_box.clear()
            self.log("Process started...")
            self.log("Please make sure both Excel files are closed before running.")

            uploader_sheet_name = "Uploader"
            target_sheet_name = "FundManagerDetails"

            self.log(f"Loading uploader file: {uploader_file}")
            uploader_wb = load_workbook(uploader_file, data_only=False)

            self.log(f"Loading FundManagerdetail file: {target_file}")
            target_wb = load_workbook(target_file)

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

                    target_cell.number_format = source_cell.number_format

                    if target_ws.max_row >= 2:
                        style_cell = target_ws.cell(
                            row=2,
                            column=target_col_num
                        )
                        self.copy_cell_style(style_cell, target_cell)

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
            self.log("Saving data directly into FundManagerdetail Excel file...")
            target_wb.save(target_file)

            uploader_wb.close()
            target_wb.close()

            self.log("FundManagerdetail Excel file updated successfully.")

            self.show_missing_data_notification(missing_data)

            self.show_success_notification(
                copied_count=copied_count,
                target_file=target_file,
                missing_count=len(missing_data)
            )

        except PermissionError:
            error_message = (
                "Permission denied while saving the FundManagerdetail file.\n\n"
                "Please close the Excel file if it is already open, then run again."
            )
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)

        except Exception as e:
            error_message = str(e)
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FundManagerCopyGUI()
    window.show()
    sys.exit(app.exec())