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


class ConsolidatePortfolioPasteGUI(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Consolidate Portfolio Data Copier")
        self.setMinimumWidth(900)

        self.client_file_path = QLineEdit()
        self.my_file_path = QLineEdit()

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        title = QLabel("Copy Portfolio Data from Client File to Your ConsolidatePortfolio File")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title)

        main_layout.addLayout(
            self.file_picker_row(
                "Client April File:",
                self.client_file_path,
                self.browse_client_file
            )
        )

        main_layout.addLayout(
            self.file_picker_row(
                "Your ConsolidatePortfolio File:",
                self.my_file_path,
                self.browse_my_file
            )
        )

        note = QLabel(
            "Note: Data will be pasted directly into your selected ConsolidatePortfolio file. No new output file will be created."
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
        label.setFixedWidth(230)

        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(browse_function)

        row.addWidget(label)
        row.addWidget(line_edit)
        row.addWidget(browse_btn)

        return row

    def browse_client_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Client April ConsolidatePortfolio File",
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
        client_file = self.client_file_path.text().strip()
        my_file = self.my_file_path.text().strip()

        if not client_file:
            QMessageBox.warning(
                self,
                "Missing File",
                "Please select the client April ConsolidatePortfolio file."
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
                "Client April file does not exist."
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

            self.log(f"Loading your file: {my_file}")
            my_wb = load_workbook(my_file)

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
                style_template[final_col] = copy(target_ws.cell(row=2, column=target_col_num)._style)
            
            self.log("Clearing old data from your target file...")
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

            self.log("Copying data from client file to your file...")

            for source_row in range(2, source_ws.max_row + 1):
                fund_value = source_ws.cell(row=source_row, column=fund_source_col).value
                security_value = source_ws.cell(row=source_row, column=security_source_col).value

                if self.is_blank(fund_value) and self.is_blank(security_value):
                    continue

                fund_name = str(fund_value).strip() if not self.is_blank(fund_value) else "Missing Fund Name"
                security_name = str(security_value).strip() if not self.is_blank(security_value) else "Missing Security Name"

                for final_col in required_columns:
                    source_col_name = resolved_source_columns[final_col]
                    target_col_name = resolved_target_columns[final_col]

                    source_col_num = source_headers[source_col_name]
                    target_col_num = target_headers[target_col_name]

                    source_cell = source_ws.cell(row=source_row, column=source_col_num)
                    target_cell = target_ws.cell(row=output_row, column=target_col_num)

                    target_cell.value = source_cell.value
                    target_cell.number_format = source_cell.number_format

                    # Apply target sheet style template
                    if final_col in style_template:
                        target_cell._style = copy(style_template[final_col])

                    # Copy number format again after style
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
            self.log("Saving data directly into your ConsolidatePortfolio file...")
            my_wb.save(my_file)

            client_wb.close()
            my_wb.close()

            self.log("Your ConsolidatePortfolio file updated successfully.")

            self.show_missing_data_notification(missing_data)

            self.show_success_notification(
                copied_count=copied_count,
                target_file=my_file,
                missing_count=len(missing_data)
            )

        except PermissionError:
            error_message = (
                "Permission denied while saving your ConsolidatePortfolio file.\n\n"
                "Please close the Excel file if it is open, then run again."
            )
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)

        except Exception as e:
            error_message = str(e)
            self.log(f"Error: {error_message}")
            self.show_error_notification(error_message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConsolidatePortfolioPasteGUI()
    window.show()
    sys.exit(app.exec())