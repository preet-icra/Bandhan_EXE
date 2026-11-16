import os
import re
import sys
import html
import traceback
import platform
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher
from typing import Any

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication, QFileDialog, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPushButton, QProgressBar, QTextEdit,
    QVBoxLayout, QWidget, QTabWidget, QGridLayout, QSplitter, QFrame, QSizePolicy
)

# ============================================================
# SECTION 2 LOGIC: Asset Allocation Excel Generator (merged from asset.py)
# ============================================================
import sys
import re
import html
import traceback
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QLabel,
    QLineEdit,
    QTextEdit,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QMessageBox,
    QSplitter,
    QGroupBox,
)


# -----------------------------
# Helper functions
# -----------------------------

def asset_normalize_text(value):
    if pd.isna(value):
        return ""

    text = html.unescape(str(value))
    text = text.strip().lower()
    text = text.replace("&", "and")
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def asset_display_text(value):
    if pd.isna(value):
        return ""
    return html.unescape(str(value)).strip()


def asset_clean_number(value):
    if pd.isna(value):
        return 0.0

    text = str(value).strip()
    text = text.replace(",", "")
    text = text.replace("%", "")

    try:
        return float(text)
    except Exception:
        return 0.0


def asset_find_fund_column(df):
    for col in df.columns:
        col_text = str(col).strip().lower()
        if "fund" in col_text:
            return col

    return df.columns[0]


def asset_find_mapping_columns(mapping_df):
    columns = list(mapping_df.columns)

    if len(columns) < 2:
        raise ValueError("Mapping file must have at least two columns.")

    actual_col = None
    consolidated_col = None

    for col in columns:
        col_l = str(col).strip().lower()
        if "actual" in col_l or "asset" in col_l:
            actual_col = col
            break

    for col in columns:
        col_l = str(col).strip().lower()
        if "consolidated" in col_l or "instrument" in col_l or "security" in col_l:
            consolidated_col = col
            break

    if actual_col is None:
        actual_col = columns[0]

    if consolidated_col is None:
        consolidated_col = columns[1]

    return actual_col, consolidated_col


def asset_is_net_equity(asset_class):
    return "netequity" in asset_normalize_text(asset_class)


def asset_is_arbitrage(asset_class):
    return "arbitrage" in asset_normalize_text(asset_class)


def asset_is_debt(asset_class):
    return "debt" in asset_normalize_text(asset_class)


def asset_is_mutual_fund(asset_class):
    key = asset_normalize_text(asset_class)
    return (
        "mutualfund" in key
        or "mutualfundunit" in key
        or "mutualfundunits" in key
    )


def asset_is_international_equity(asset_class):
    key = asset_normalize_text(asset_class)
    return "internationalequity" in key or "foreign" in key or "overseas" in key


def asset_is_gold_silver(asset_class):
    key = asset_normalize_text(asset_class)
    return "gold" in key or "silver" in key


def asset_is_gold_or_silver_text(value):
    key = asset_normalize_text(value)
    return "gold" in key or "silver" in key


def asset_is_exchange_traded_fund_name(value):
    key = asset_normalize_text(value)
    return (
        "exchangetradedfund" in key
        or "exchangetradedfunds" in key
        or "etf" in key
    )


def asset_is_international_etf_name(value):
    key = asset_normalize_text(value)
    return (
        ("international" in key or "foreign" in key or "overseas" in key)
        and asset_is_exchange_traded_fund_name(value)
    )


def asset_is_derivative_or_preference(name):
    key = asset_normalize_text(name)
    return (
        "future" in key
        or "futures" in key
        or "option" in key
        or "options" in key
        or "preferenceshare" in key
        or "preferenceshares" in key
    )


def asset_make_net_equity_expression(net_equity_parts):
    """
    Net Equity formula logic.

    If derivative/preference value is negative:
        B6--B7

    If derivative/preference value is positive:
        B6+B7
    """
    expression = ""

    for part in net_equity_parts:
        cell_ref = part["cell"]
        raw_value = float(part["value"])
        is_subtract_item = part["is_subtract_item"]

        if expression == "":
            expression = cell_ref
            continue

        if is_subtract_item:
            if raw_value < 0:
                expression += f"--{cell_ref}"
            else:
                expression += f"+{cell_ref}"
        else:
            expression += f"+{cell_ref}"

    if expression == "":
        expression = "0"

    return expression


def asset_parse_fund_names(text):
    fund_names = []

    for line in text.splitlines():
        name = line.strip()
        if name:
            fund_names.append(name)

    return fund_names


def asset_add_report(report_rows, log_callback, fund_name, asset_class, item_name, issue_type, source_column):
    row = {
        "Fund Name": fund_name,
        "Asset Class": asset_class,
        "Item Name": item_name,
        "Issue Type": issue_type,
        "Source Column": source_column,
    }

    report_rows.append(row)

    log_callback(
        f"Report | Fund: {fund_name} | Asset Class: {asset_class} | "
        f"Item: {item_name} | Issue: {issue_type} | Column: {source_column}"
    )


def asset_get_non_gold_silver_non_international_etf_value(fund_df):
    """
    Mutual Fund Units ETF logic.

    Add Exchange Traded Funds to Mutual Fund Units only when:
    1. Instrument Type is ETF / Exchange Traded Funds
    2. Security Name does NOT contain Gold or Silver
    3. Instrument Type is NOT International Exchange Traded Funds

    This prevents Bandhan Multi Asset Allocation Fund from showing Mutual Fund Units
    when ETF values are Gold/Silver or International ETF.
    """
    etf_mask = fund_df["_instrument_display"].apply(asset_is_exchange_traded_fund_name)
    gold_silver_security_mask = fund_df["_security_display"].apply(asset_is_gold_or_silver_text)
    international_etf_mask = fund_df["_instrument_display"].apply(asset_is_international_etf_name)

    final_mask = etf_mask & (~gold_silver_security_mask) & (~international_etf_mask)

    return float(fund_df.loc[final_mask, "_percent_value"].sum())


# -----------------------------
# Mapping report checks
# -----------------------------

def asset_check_instrument_mapping_report(
    fund_name,
    asset_order,
    asset_to_names,
    mapping_key_to_display,
    row_by_instrument_key,
    report_rows,
    log_callback,
):
    all_mapping_keys = set()

    for asset in asset_order:
        mapped_names = asset_to_names.get(asset, [])

        if not mapped_names and not asset_is_gold_silver(asset):
            asset_add_report(
                report_rows,
                log_callback,
                fund_name,
                asset,
                "",
                "Mapping Consolidated Name is blank or missing for this Asset Class",
                "Mapping File - Consolidated Name",
            )
            continue

        for mapped_name in mapped_names:
            mapped_key = asset_normalize_text(mapped_name)

            if mapped_key == "":
                asset_add_report(
                    report_rows,
                    log_callback,
                    fund_name,
                    asset,
                    mapped_name,
                    "Mapping Consolidated Name is blank",
                    "Mapping File - Consolidated Name",
                )
                continue

            all_mapping_keys.add(mapped_key)

            if mapped_key not in row_by_instrument_key:
                asset_add_report(
                    report_rows,
                    log_callback,
                    fund_name,
                    asset,
                    mapped_name,
                    "Mapping Instrument Type missing in Consolidate Instrument Type",
                    "Consolidate File - Column E Instrument Type",
                )

    for instrument_key in row_by_instrument_key.keys():
        if instrument_key == "":
            continue

        if instrument_key not in all_mapping_keys:
            instrument_display = mapping_key_to_display.get(instrument_key, instrument_key)

            if asset_is_exchange_traded_fund_name(instrument_display):
                asset_add_report(
                    report_rows,
                    log_callback,
                    fund_name,
                    "Mutual Fund Units / International Equity / Gold / Silver",
                    instrument_display,
                    "Exchange Traded Funds found in Consolidate. Check mapping according to ETF type.",
                    "Mapping File - Consolidated Name",
                )
                continue

            asset_add_report(
                report_rows,
                log_callback,
                fund_name,
                "Not mapped",
                instrument_display,
                "Consolidate Instrument Type missing in Mapping",
                "Mapping File - Consolidated Name",
            )


# -----------------------------
# Excel fund block generation
# -----------------------------

def asset_write_single_fund_block(
    ws,
    fund_df,
    fund_name,
    asset_order,
    asset_to_names,
    start_row,
    styles,
    log_callback,
    report_rows,
):
    yellow_fill = styles["yellow_fill"]
    header_fill = styles["header_fill"]
    border = styles["border"]
    bold_font = styles["bold_font"]
    center = styles["center"]
    left = styles["left"]

    log_callback(f"Creating block for fund: {fund_name}")

    grouped = (
        fund_df
        .groupby("_instrument_key", as_index=False)
        .agg(
            Instrument_Type=("_instrument_display", "first"),
            Percent_Sum=("_percent_value", "sum"),
        )
    )

    grouped = grouped[grouped["_instrument_key"] != ""]
    grouped = grouped.sort_values("Percent_Sum", ascending=False).reset_index(drop=True)

    instrument_lookup = dict(zip(grouped["_instrument_key"], grouped["Percent_Sum"]))
    mapping_key_to_display = dict(zip(grouped["_instrument_key"], grouped["Instrument_Type"]))

    # -----------------------------
    # Left block
    # -----------------------------
    ws.cell(start_row, 1).value = "Fund Name"
    ws.cell(start_row, 2).value = fund_name

    ws.cell(start_row + 1, 1).value = "Row Labels"
    ws.cell(start_row + 1, 2).value = "Sum of % to Net Assets"

    ws.cell(start_row, 1).font = bold_font
    ws.cell(start_row + 1, 1).font = bold_font
    ws.cell(start_row + 1, 2).font = bold_font

    ws.cell(start_row + 1, 1).fill = header_fill
    ws.cell(start_row + 1, 2).fill = header_fill

    data_start_row = start_row + 2
    row_by_instrument_key = {}

    for i, row in grouped.iterrows():
        excel_row = data_start_row + i

        instrument_name = row["Instrument_Type"]
        percent_sum = float(row["Percent_Sum"])
        instrument_key = row["_instrument_key"]

        ws.cell(excel_row, 1).value = instrument_name
        ws.cell(excel_row, 2).value = percent_sum
        ws.cell(excel_row, 2).number_format = "0.0000"

        row_by_instrument_key[instrument_key] = excel_row

    total_row = data_start_row + len(grouped)

    ws.cell(total_row, 1).value = "Grand Total"
    ws.cell(total_row, 2).value = f"=SUM(B{data_start_row}:B{total_row - 1})"
    ws.cell(total_row, 1).font = bold_font
    ws.cell(total_row, 2).font = bold_font
    ws.cell(total_row, 2).number_format = "0.0000"

    for row_num in range(start_row, total_row + 1):
        for col_num in range(1, 3):
            ws.cell(row_num, col_num).border = border

    asset_check_instrument_mapping_report(
        fund_name=fund_name,
        asset_order=asset_order,
        asset_to_names=asset_to_names,
        mapping_key_to_display=mapping_key_to_display,
        row_by_instrument_key=row_by_instrument_key,
        report_rows=report_rows,
        log_callback=log_callback,
    )

    # -----------------------------
    # Right yellow block
    # -----------------------------
    ws.cell(start_row + 1, 4).value = "Asset class"
    ws.cell(start_row + 1, 5).value = "%"
    ws.cell(start_row + 1, 6).value = "Descending order in all fund"

    for col_num in range(4, 7):
        cell = ws.cell(start_row + 1, col_num)
        cell.fill = yellow_fill
        cell.font = bold_font
        cell.border = border
        cell.alignment = center

    asset_row = start_row + 2
    asset_rows_written = 0

    for asset in asset_order:
        mapped_names = asset_to_names.get(asset, [])

        formula = None
        has_any_available_name = False

        # ----------------------------------------------------
        # Gold/Silver logic:
        # Value comes from Security Name column C.
        # ----------------------------------------------------
        if asset_is_gold_silver(asset):
            asset_key = asset_normalize_text(asset)

            keywords = []

            if "gold" in asset_key:
                keywords.append("gold")

            if "silver" in asset_key:
                keywords.append("silver")

            if not keywords:
                keywords = ["gold", "silver"]

            mask = pd.Series(False, index=fund_df.index)

            for keyword in keywords:
                mask = mask | fund_df["_security_key"].str.contains(
                    keyword,
                    case=False,
                    na=False,
                )

            gold_silver_value = float(fund_df.loc[mask, "_percent_value"].sum())

            if gold_silver_value != 0:
                has_any_available_name = True
                formula = f"={gold_silver_value}/100"

        else:
            sum_items = []
            net_equity_parts = []

            for mapped_name in mapped_names:
                mapped_key = asset_normalize_text(mapped_name)

                # For Mutual Fund Units + Exchange Traded Funds:
                # Do not directly use grouped B row because it may include
                # Gold/Silver ETF or International ETF.
                # ETF is handled separately below using raw data.
                if asset_is_mutual_fund(asset) and asset_is_exchange_traded_fund_name(mapped_name):
                    continue

                if mapped_key not in row_by_instrument_key:
                    continue

                has_any_available_name = True

                raw_value = float(instrument_lookup.get(mapped_key, 0.0))
                excel_ref = f"B{row_by_instrument_key[mapped_key]}"

                if asset_is_net_equity(asset):
                    subtract_item = asset_is_derivative_or_preference(mapped_name)

                    net_equity_parts.append(
                        {
                            "cell": excel_ref,
                            "value": raw_value,
                            "is_subtract_item": subtract_item,
                        }
                    )

                elif asset_is_arbitrage(asset):
                    sum_items.append(excel_ref)

                elif asset_is_debt(asset):
                    sum_items.append(excel_ref)

                elif asset_is_mutual_fund(asset) or asset_is_international_equity(asset):
                    sum_items.append(excel_ref)

                else:
                    sum_items.append(excel_ref)

            # ----------------------------------------------------
            # Mutual Fund Units logic:
            #
            # Add ETF to Mutual Fund Units only if ETF is:
            # - not Gold/Silver security
            # - not International ETF
            #
            # This removes Mutual Fund Units from Bandhan Multi Asset Allocation Fund.
            # ----------------------------------------------------
            if asset_is_mutual_fund(asset):
                eligible_etf_value = asset_get_non_gold_silver_non_international_etf_value(fund_df)

                if eligible_etf_value != 0:
                    sum_items.append(str(eligible_etf_value))
                    has_any_available_name = True

                    log_callback(
                        f"Exchange Traded Funds added under Mutual Fund Units "
                        f"because Security Name is not Gold/Silver and ETF is not International | "
                        f"Fund: {fund_name} | Value: {eligible_etf_value}"
                    )

            if not has_any_available_name:
                formula = None

            elif asset_is_net_equity(asset):
                expression = asset_make_net_equity_expression(net_equity_parts)
                formula = f"=({expression})/100"

            elif asset_is_arbitrage(asset):
                if sum_items:
                    formula = f"=-SUM({','.join(sum_items)})/100"
                else:
                    formula = None

            else:
                if sum_items:
                    formula = f"=SUM({','.join(sum_items)})/100"
                else:
                    formula = None

        if not has_any_available_name or formula is None:
            continue

        current_row = asset_row + asset_rows_written

        ws.cell(current_row, 4).value = asset
        ws.cell(current_row, 5).value = formula
        ws.cell(current_row, 5).number_format = "0.00%"

        ws.cell(current_row, 6).value = f"=E{current_row}*100"
        ws.cell(current_row, 6).number_format = "0.00"

        for col_num in range(4, 7):
            cell = ws.cell(current_row, col_num)
            cell.fill = yellow_fill
            cell.border = border
            cell.alignment = left if col_num == 4 else center

        asset_rows_written += 1

    # Total row for asset block
    asset_total_row = asset_row + asset_rows_written + 2

    ws.cell(asset_total_row, 4).value = "Total"

    if asset_rows_written > 0:
        ws.cell(asset_total_row, 5).value = f"=SUM(E{asset_row}:E{asset_row + asset_rows_written - 1})"
    else:
        ws.cell(asset_total_row, 5).value = 0

    ws.cell(asset_total_row, 5).number_format = "0.00%"

    for col_num in range(4, 6):
        cell = ws.cell(asset_total_row, col_num)
        cell.fill = yellow_fill
        cell.border = border
        cell.alignment = left if col_num == 4 else center

    for row_num in range(start_row, asset_total_row + 1):
        for col_num in range(4, 7):
            ws.cell(row_num, col_num).fill = yellow_fill

    section_end_row = max(total_row, asset_total_row)

    return section_end_row + 2


# -----------------------------
# Report sheet
# -----------------------------

def asset_create_report_sheet(wb, report_rows):
    ws = wb.create_sheet("Mapping Report")

    headers = [
        "Fund Name",
        "Asset Class",
        "Item Name",
        "Issue Type",
        "Source Column",
    ]

    header_fill = PatternFill("solid", fgColor="FFC000")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold_font = Font(bold=True)

    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(1, col_index)
        cell.value = header
        cell.font = bold_font
        cell.fill = header_fill
        cell.border = border

    if report_rows:
        for row_index, item in enumerate(report_rows, start=2):
            ws.cell(row_index, 1).value = item.get("Fund Name", "")
            ws.cell(row_index, 2).value = item.get("Asset Class", "")
            ws.cell(row_index, 3).value = item.get("Item Name", "")
            ws.cell(row_index, 4).value = item.get("Issue Type", "")
            ws.cell(row_index, 5).value = item.get("Source Column", "")

            for col_index in range(1, 6):
                ws.cell(row_index, col_index).border = border
    else:
        ws.cell(2, 1).value = "No mapping mismatch found."

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 75
    ws.column_dimensions["E"].width = 45

    ws.freeze_panes = "A2"


# -----------------------------
# Main Excel generation logic
# -----------------------------

def asset_generate_output_excel(consolidate_path, mapping_path, output_path, fund_names, log_callback):
    report_rows = []

    log_callback("Reading Consolidate file...")

    consolidate_df = pd.read_excel(
        consolidate_path,
        sheet_name=0,
        engine="openpyxl",
        dtype=object,
    )

    if consolidate_df.shape[1] < 5:
        raise ValueError(
            "Consolidate file must have at least 5 columns because Column E Instrument Type is required."
        )

    percent_col = consolidate_df.columns[1]
    security_col = consolidate_df.columns[2]
    instrument_col = consolidate_df.columns[4]
    fund_col = asset_find_fund_column(consolidate_df)

    log_callback(f"Fund column detected: {fund_col}")
    log_callback(f"% to Net Assets column used: {percent_col}")
    log_callback(f"Security Name column used: {security_col}")
    log_callback(f"Instrument Type column used: {instrument_col}")

    consolidate_df["_fund_text"] = consolidate_df[fund_col].astype(str).str.strip()
    consolidate_df["_percent_value"] = consolidate_df[percent_col].apply(asset_clean_number)
    consolidate_df["_security_display"] = consolidate_df[security_col].apply(asset_display_text)
    consolidate_df["_instrument_display"] = consolidate_df[instrument_col].apply(asset_display_text)
    consolidate_df["_instrument_key"] = consolidate_df["_instrument_display"].apply(asset_normalize_text)
    consolidate_df["_security_key"] = consolidate_df["_security_display"].apply(asset_normalize_text)

    # -----------------------------
    # Read mapping file
    # -----------------------------
    log_callback("Reading Mapping file...")

    mapping_df = pd.read_excel(
        mapping_path,
        sheet_name=0,
        engine="openpyxl",
        dtype=object,
    )

    if mapping_df.shape[1] < 2:
        raise ValueError("Mapping file must have at least two columns: Actual Name and Consolidated Name.")

    actual_col, consolidated_col = asset_find_mapping_columns(mapping_df)

    log_callback(f"Mapping Actual Name column used: {actual_col}")
    log_callback(f"Mapping Consolidated Name column used: {consolidated_col}")

    mapping_df = mapping_df[[actual_col, consolidated_col]].copy()
    mapping_df.columns = ["Actual_Name", "Consolidated_Name"]

    mapping_df = mapping_df.dropna(subset=["Actual_Name"])

    mapping_df["Actual_Name"] = mapping_df["Actual_Name"].apply(asset_display_text)
    mapping_df["Consolidated_Name"] = mapping_df["Consolidated_Name"].apply(asset_display_text)

    asset_order = []
    asset_to_names = {}

    for _, row in mapping_df.iterrows():
        asset = row["Actual_Name"]
        consolidated_name = row["Consolidated_Name"]

        if asset not in asset_to_names:
            asset_to_names[asset] = []
            asset_order.append(asset)

        if consolidated_name != "":
            asset_to_names[asset].append(consolidated_name)

    if not asset_order:
        raise ValueError("No asset classes found in mapping file.")

    # Ensure Mutual Fund Units asset class exists for ETF-only funds.
    # It will appear only when eligible non-Gold/Silver and non-International ETF value exists.
    mutual_fund_asset_exists = any(asset_is_mutual_fund(asset) for asset in asset_order)

    if not mutual_fund_asset_exists:
        asset_order.append("Mutual Fund Units")
        asset_to_names["Mutual Fund Units"] = []
        log_callback("Mutual Fund Units asset class added automatically for eligible Exchange Traded Funds logic.")

    # -----------------------------
    # Create workbook
    # -----------------------------
    log_callback("Creating output Excel file...")

    wb = Workbook()
    ws = wb.active
    ws.title = "Output"

    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    header_fill = PatternFill("solid", fgColor="D9EAF7")

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    bold_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    styles = {
        "yellow_fill": yellow_fill,
        "header_fill": header_fill,
        "border": border,
        "bold_font": bold_font,
        "center": center,
        "left": left,
    }

    current_start_row = 1
    processed_count = 0

    for fund_name in fund_names:
        fund_mask = consolidate_df["_fund_text"].str.contains(
            str(fund_name).strip(),
            case=False,
            na=False,
            regex=False,
        )

        fund_df = consolidate_df.loc[fund_mask].copy()

        if fund_df.empty:
            asset_add_report(
                report_rows,
                log_callback,
                fund_name,
                "All",
                fund_name,
                "Fund name not found in Consolidate file",
                "Consolidate File - Fund Name Column",
            )
            continue

        log_callback(f"Rows found for {fund_name}: {len(fund_df)}")

        current_start_row = asset_write_single_fund_block(
            ws=ws,
            fund_df=fund_df,
            fund_name=fund_name,
            asset_order=asset_order,
            asset_to_names=asset_to_names,
            start_row=current_start_row,
            styles=styles,
            log_callback=log_callback,
            report_rows=report_rows,
        )

        processed_count += 1

    if processed_count == 0:
        asset_create_report_sheet(wb, report_rows)
        output_path = Path(output_path)
        wb.save(output_path)
        raise ValueError(
            "No matching fund data found for the given fund names. "
            "Mapping Report sheet has been created in the output file."
        )

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 32
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 28

    ws.freeze_panes = "A3"

    asset_create_report_sheet(wb, report_rows)

    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    output_path = Path(output_path)
    wb.save(output_path)

    log_callback(f"Output file saved successfully: {output_path}")

    if report_rows:
        warning_text = (
            f"Output generated with mapping report warnings.\n\n"
            f"Total report items: {len(report_rows)}\n"
            f"Please check the 'Mapping Report' sheet in output Excel."
        )
    else:
        warning_text = "Output generated successfully. No mapping mismatch found."

    return str(output_path), warning_text


# -----------------------------
# Worker thread
# -----------------------------

class asset_ExcelWorker(QThread):
    log = pyqtSignal(str)
    success = pyqtSignal(str, str)
    error = pyqtSignal(str)

    def __init__(self, consolidate_path, mapping_path, output_path, fund_names):
        super().__init__()
        self.consolidate_path = consolidate_path
        self.mapping_path = mapping_path
        self.output_path = output_path
        self.fund_names = fund_names

    def run(self):
        try:
            output_path, warning_text = asset_generate_output_excel(
                self.consolidate_path,
                self.mapping_path,
                self.output_path,
                self.fund_names,
                self.log.emit,
            )
            self.success.emit(output_path, warning_text)
        except Exception as e:
            details = traceback.format_exc()
            self.error.emit(f"{str(e)}\n\n{details}")


# -----------------------------
# PyQt6 GUI
# -----------------------------

class asset_FundExcelGenerator(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fund Excel Output Generator")
        self.resize(1150, 700)

        self.worker = None
        self.build_ui()

    def build_ui(self):
        splitter = QSplitter()

        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout()

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        log_layout.addWidget(self.log_box)
        log_group.setLayout(log_layout)

        process_group = QGroupBox("Process")
        process_layout = QVBoxLayout()

        grid = QGridLayout()

        self.consolidate_input = QLineEdit()
        self.mapping_input = QLineEdit()
        self.output_input = QLineEdit()

        self.fund_input = QTextEdit()
        self.fund_input.setPlaceholderText(
            "Enter one fund name per line.\n"
            "Example:\n"
            "Bandhan Arbitrage Fund\n"
            "Bandhan Equity Savings Fund\n"
            "Bandhan Conservative Hybrid Fund\n"
            "Bandhan Multi Asset Allocation Fund"
        )

        btn_consolidate = QPushButton("Browse")
        btn_mapping = QPushButton("Browse")
        btn_output = QPushButton("Save As")

        btn_consolidate.clicked.connect(self.browse_consolidate)
        btn_mapping.clicked.connect(self.browse_mapping)
        btn_output.clicked.connect(self.browse_output)

        grid.addWidget(QLabel("Consolidate File:"), 0, 0)
        grid.addWidget(self.consolidate_input, 0, 1)
        grid.addWidget(btn_consolidate, 0, 2)

        grid.addWidget(QLabel("Mapping File:"), 1, 0)
        grid.addWidget(self.mapping_input, 1, 1)
        grid.addWidget(btn_mapping, 1, 2)

        grid.addWidget(QLabel("Output File:"), 2, 0)
        grid.addWidget(self.output_input, 2, 1)
        grid.addWidget(btn_output, 2, 2)

        grid.addWidget(QLabel("Fund Names:"), 3, 0)
        grid.addWidget(self.fund_input, 3, 1, 1, 2)

        self.process_button = QPushButton("Generate Output Excel")
        self.process_button.clicked.connect(self.start_process)

        process_layout.addLayout(grid)
        process_layout.addWidget(self.process_button)
        process_layout.addStretch()

        process_group.setLayout(process_layout)

        splitter.addWidget(log_group)
        splitter.addWidget(process_group)
        splitter.setSizes([520, 630])

        main_layout = QHBoxLayout()
        main_layout.addWidget(splitter)

        self.setLayout(main_layout)

    def log(self, message):
        self.log_box.append(message)

    def browse_consolidate(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Consolidate Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls)",
        )

        if file_path:
            self.consolidate_input.setText(file_path)

    def browse_mapping(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Mapping Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls)",
        )

        if file_path:
            self.mapping_input.setText(file_path)

    def browse_output(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output Excel File",
            "output.xlsx",
            "Excel Files (*.xlsx)",
        )

        if file_path:
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"

            self.output_input.setText(file_path)

    def start_process(self):
        consolidate_path = self.consolidate_input.text().strip()
        mapping_path = self.mapping_input.text().strip()
        output_path = self.output_input.text().strip()
        fund_names_text = self.fund_input.toPlainText().strip()

        fund_names = asset_parse_fund_names(fund_names_text)

        if not consolidate_path:
            QMessageBox.warning(self, "Missing File", "Please select Consolidate file.")
            return

        if not mapping_path:
            QMessageBox.warning(self, "Missing File", "Please select Mapping file.")
            return

        if not output_path:
            QMessageBox.warning(self, "Missing Output", "Please select output file path.")
            return

        if not fund_names:
            QMessageBox.warning(self, "Missing Fund Names", "Please enter at least one fund name.")
            return

        self.log_box.clear()
        self.log("Starting process...")
        self.log(f"Total funds entered: {len(fund_names)}")

        self.process_button.setEnabled(False)

        self.worker = asset_ExcelWorker(
            consolidate_path,
            mapping_path,
            output_path,
            fund_names,
        )

        self.worker.log.connect(self.log)
        self.worker.success.connect(self.on_success)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_success(self, output_path, warning_text):
        self.process_button.setEnabled(True)
        self.log("Process completed.")
        self.log(warning_text)

        if "warnings" in warning_text.lower():
            QMessageBox.warning(
                self,
                "Completed with Mapping Report",
                f"{warning_text}\n\nOutput file:\n{output_path}",
            )
        else:
            QMessageBox.information(
                self,
                "Success",
                f"{warning_text}\n\nOutput file:\n{output_path}",
            )

    def on_error(self, error_message):
        self.process_button.setEnabled(True)
        self.log("Error occurred.")
        self.log(error_message)

        QMessageBox.critical(
            self,
            "Error",
            error_message,
        )


if False and __name__ == "__main__":
    app = QApplication(sys.argv)
    window = asset_FundExcelGenerator()
    window.show()
    sys.exit(app.exec())


# ============================================================
# SECTION 3 LOGIC: Hybrid Asset Allocation Snapshot Update (merged from hybridsnapshot.py)
# ============================================================
import os
import sys
import re
import traceback
import platform

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QPushButton,
    QFileDialog,
    QLineEdit,
    QTextEdit,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QMessageBox,
    QFrame,
    QSizePolicy
)

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Pt


# ============================================================
# CONFIGURATION
# ============================================================

PPT_TABLE_OBJECT_NAME = "SnapshotTab"

EXCEL_FUND_LABEL = "Fund Name"
EXCEL_ASSET_CLASS_HEADER = "Asset class"
EXCEL_PERCENT_HEADER = "%"

PPT_SCHEME_HEADER = "Scheme Name"

PPT_ASSET_ALLOCATION_HEADER_VARIANTS = [
    "Asset Allocation(%)",
    "Asset Allocation (%)",
    "Asset Allocation %",
    "Asset Allocation"
]

PPT_FONT_NAME = "Ubuntu"
PPT_FONT_SIZE = 6

USE_EXCEL_COM_ON_WINDOWS = True


# ============================================================
# COMMON HELPER FUNCTIONS
# ============================================================

def hybrid_normalize_text(value):
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\n", " ")
    text = " ".join(text.split())
    return text.strip().lower()


def hybrid_normalize_compact(value):
    return hybrid_normalize_text(value).replace(" ", "")


def hybrid_normalize_for_fund_match(value):
    """
    Ignores symbols, spaces, brackets, punctuation and special characters.
    Keeps only letters and numbers.
    """
    if value is None:
        return ""

    text = str(value).lower()
    text = text.replace("\n", " ")
    text = re.sub(r"[^a-z0-9]", "", text)
    return text.strip()


def hybrid_clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def hybrid_format_percentage(value):
    """
    Formats Excel value.

    Handles:
    0.713       -> 71.30%
    71.3        -> 71.30%
    '71.30%'    -> 71.30%
    '' / None   -> ''
    """
    if value is None:
        return ""

    if isinstance(value, str):
        value = value.strip()

        if value == "":
            return ""

        if value.endswith("%"):
            return value

        try:
            number = float(value.replace(",", ""))
        except ValueError:
            return value
    else:
        try:
            number = float(value)
        except Exception:
            return str(value)

    if abs(number) <= 1:
        number = number * 100

    return f"{number:.2f}%"


def hybrid_iter_shapes(shapes):
    for shape in shapes:
        yield shape

        if hasattr(shape, "shapes"):
            for sub_shape in hybrid_iter_shapes(shape.shapes):
                yield sub_shape


def hybrid_set_ppt_cell_text(cell, text, font_name=PPT_FONT_NAME, font_size=PPT_FONT_SIZE):
    """
    Sets text into PPT table cell and applies Ubuntu font size 6.
    """
    cell.text = ""

    text_frame = cell.text_frame
    text_frame.clear()

    lines = text.split("\n")

    for index, line in enumerate(lines):
        if index == 0:
            paragraph = text_frame.paragraphs[0]
        else:
            paragraph = text_frame.add_paragraph()

        paragraph.text = ""

        run = paragraph.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)


# ============================================================
# EXCEL VALUE READER
# ============================================================

class hybrid_ExcelValueReader:
    """
    Reads Excel calculated values.

    Priority:
    1. Excel COM on Windows, if available.
       This reads real calculated values from formulas.
    2. openpyxl data_only=True fallback.
       This reads cached calculated values only.
    """

    def __init__(self, excel_path, log):
        self.excel_path = excel_path
        self.log = log

        self.mode = "openpyxl"
        self.workbook = None
        self.sheet = None

        self.excel_app = None
        self.com_workbook = None
        self.com_sheet = None

        self.max_row = 1

    def open(self):
        if USE_EXCEL_COM_ON_WINDOWS and platform.system().lower() == "windows":
            try:
                self._open_with_com()
                return
            except Exception as e:
                self.log(f"Excel COM reading failed. Falling back to openpyxl. Reason: {e}")

        self._open_with_openpyxl()

    def _open_with_com(self):
        self.log("Opening Excel using Microsoft Excel COM for calculated formula values...")

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()

        self.excel_app = win32com.client.DispatchEx("Excel.Application")
        self.excel_app.Visible = False
        self.excel_app.DisplayAlerts = False

        self.com_workbook = self.excel_app.Workbooks.Open(os.path.abspath(self.excel_path))
        self.excel_app.CalculateFullRebuild()

        self.com_sheet = self.com_workbook.Worksheets(1)
        self.max_row = self.com_sheet.UsedRange.Rows.Count

        self.mode = "com"

        self.log("Excel COM mode active. Formula calculated values will be read.")

    def _open_with_openpyxl(self):
        self.log("Opening Excel using openpyxl data_only=True...")
        self.log("Note: openpyxl reads saved cached formula values only.")

        self.workbook = load_workbook(self.excel_path, data_only=True)
        self.sheet = self.workbook.active
        self.max_row = self.sheet.max_row

        self.mode = "openpyxl"

    def value(self, row, col):
        if self.mode == "com":
            return self.com_sheet.Cells(row, col).Value

        return self.sheet.cell(row=row, column=col).value

    def close(self):
        if self.mode == "com":
            try:
                if self.com_workbook is not None:
                    self.com_workbook.Close(SaveChanges=False)
            except Exception:
                pass

            try:
                if self.excel_app is not None:
                    self.excel_app.Quit()
            except Exception:
                pass

            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

        else:
            try:
                if self.workbook is not None:
                    self.workbook.close()
            except Exception:
                pass


# ============================================================
# EXCEL READER
# ============================================================

def hybrid_read_excel_asset_allocations(excel_path, log):
    """
    Reads Excel fund-wise Asset Allocation data.

    Expected layout:
    Column A: Fund Name label
    Column B: Actual Fund Name

    Column D: Asset class
    Column E: %

    It reads values only, not formulas.
    """

    reader = hybrid_ExcelValueReader(excel_path, log)
    reader.open()

    excel_data = {}

    try:
        max_row = reader.max_row
        row = 1

        while row <= max_row:
            col_a_value = hybrid_normalize_text(reader.value(row, 1))

            if col_a_value == hybrid_normalize_text(EXCEL_FUND_LABEL):
                fund_name = hybrid_clean_text(reader.value(row, 2))

                if not fund_name:
                    log(f"Row {row}: Fund Name label found but fund value is blank. Skipping.")
                    row += 1
                    continue

                fund_match_key = hybrid_normalize_for_fund_match(fund_name)

                log(f"Found Excel fund: {fund_name}")

                header_row = None
                search_row = row

                while search_row <= max_row:
                    next_fund_check = hybrid_normalize_text(reader.value(search_row, 1))

                    if search_row != row and next_fund_check == hybrid_normalize_text(EXCEL_FUND_LABEL):
                        break

                    col_d_value = hybrid_normalize_text(reader.value(search_row, 4))
                    col_e_value = hybrid_normalize_text(reader.value(search_row, 5))

                    if (
                        col_d_value == hybrid_normalize_text(EXCEL_ASSET_CLASS_HEADER)
                        and col_e_value == hybrid_normalize_text(EXCEL_PERCENT_HEADER)
                    ):
                        header_row = search_row
                        break

                    search_row += 1

                if header_row is None:
                    log(f"Asset class / % header not found for fund: {fund_name}")
                    row += 1
                    continue

                allocations = []
                data_row = header_row + 1

                while data_row <= max_row:
                    asset_class_value = reader.value(data_row, 4)
                    percentage_value = reader.value(data_row, 5)

                    asset_class_norm = hybrid_normalize_text(asset_class_value)

                    if asset_class_norm == "total":
                        break

                    if hybrid_normalize_text(reader.value(data_row, 1)) == hybrid_normalize_text(EXCEL_FUND_LABEL):
                        break

                    if asset_class_value is None and percentage_value is None:
                        data_row += 1
                        continue

                    if asset_class_value is not None:
                        asset_class_text = hybrid_clean_text(asset_class_value)
                        percentage_text = hybrid_format_percentage(percentage_value)

                        log(f"  Excel allocation: {asset_class_text} = {percentage_text}")

                        allocations.append((asset_class_text, percentage_text))

                    data_row += 1

                if allocations:
                    excel_data[fund_match_key] = {
                        "fund_name": fund_name,
                        "allocations": allocations
                    }

                    log(f"Loaded {len(allocations)} allocation rows for: {fund_name}")
                else:
                    log(f"No allocation rows found for fund: {fund_name}")

                row = data_row

            row += 1

    finally:
        reader.close()

    log(f"Excel reading completed. Total funds loaded: {len(excel_data)}")
    return excel_data


# ============================================================
# POWERPOINT PROCESSOR
# ============================================================

def hybrid_find_all_snapshot_tables(prs, log):
    """
    Finds all PowerPoint table objects named SnapshotTab across every slide.
    """
    tables = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        for shape in hybrid_iter_shapes(slide.shapes):
            shape_name = getattr(shape, "name", "")

            if shape_name == PPT_TABLE_OBJECT_NAME and getattr(shape, "has_table", False):
                log(f"Found table '{PPT_TABLE_OBJECT_NAME}' on slide {slide_index}.")
                tables.append({
                    "slide_index": slide_index,
                    "table": shape.table
                })

    return tables


def hybrid_find_table_columns(table, log, slide_index):
    scheme_col = None
    asset_allocation_col = None
    header_row = None

    asset_header_variants = [
        hybrid_normalize_compact(item)
        for item in PPT_ASSET_ALLOCATION_HEADER_VARIANTS
    ]

    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            cell_text = cell.text
            normalized = hybrid_normalize_text(cell_text)
            compact = hybrid_normalize_compact(cell_text)

            if normalized == hybrid_normalize_text(PPT_SCHEME_HEADER):
                scheme_col = c_idx
                header_row = r_idx

            if compact in asset_header_variants:
                asset_allocation_col = c_idx
                header_row = r_idx

        if scheme_col is not None and asset_allocation_col is not None:
            break

    if scheme_col is None:
        raise Exception(f"Could not find 'Scheme Name' column in SnapshotTab on slide {slide_index}.")

    if asset_allocation_col is None:
        raise Exception(f"Could not find 'Asset Allocation(%)' column in SnapshotTab on slide {slide_index}.")

    log(f"Slide {slide_index}: Scheme Name column = {scheme_col + 1}")
    log(f"Slide {slide_index}: Asset Allocation column = {asset_allocation_col + 1}")

    return header_row, scheme_col, asset_allocation_col


def hybrid_find_matching_excel_fund(scheme_name, excel_data):
    ppt_scheme_key = hybrid_normalize_for_fund_match(scheme_name)

    if ppt_scheme_key in excel_data:
        return ppt_scheme_key

    return None


def hybrid_build_asset_allocation_text(allocations):
    """
    Builds final PPT text.

    Example:
    Net Equity - -0.67%
    Arbitrage - 71.30%
    Debt - 4.11%
    Mutual fund Units - 25.26%
    """
    lines = []

    for asset_class, percentage in allocations:
        if percentage:
            lines.append(f"{asset_class} : {percentage}")
        else:
            lines.append(f"{asset_class} :")

    return "\n".join(lines)


def hybrid_update_ppt_from_excel(ppt_path, excel_data, output_path, log):
    log("Opening PowerPoint file...")

    prs = Presentation(ppt_path)

    snapshot_tables = hybrid_find_all_snapshot_tables(prs, log)

    if not snapshot_tables:
        raise Exception(f"No PPT table object named '{PPT_TABLE_OBJECT_NAME}' found in any slide.")

    total_updated_count = 0
    not_found = []

    for table_info in snapshot_tables:
        slide_index = table_info["slide_index"]
        table = table_info["table"]

        log("")
        log(f"Processing SnapshotTab on slide {slide_index}...")

        header_row, scheme_col, asset_allocation_col = hybrid_find_table_columns(
            table,
            log,
            slide_index
        )

        for r_idx in range(header_row + 1, len(table.rows)):
            scheme_cell = table.cell(r_idx, scheme_col)
            scheme_name = hybrid_clean_text(scheme_cell.text)

            if not scheme_name:
                continue

            excel_key = hybrid_find_matching_excel_fund(scheme_name, excel_data)

            if excel_key is None:
                not_found.append(f"Slide {slide_index}: {scheme_name}")
                log(f"Slide {slide_index}: No Excel match for PPT Scheme Name: {scheme_name}")
                continue

            allocations = excel_data[excel_key]["allocations"]
            paste_text = hybrid_build_asset_allocation_text(allocations)

            log(f"Slide {slide_index}: Pasting for {scheme_name}:")
            for asset_class, percentage in allocations:
                log(f"  {asset_class} - {percentage}")

            target_cell = table.cell(r_idx, asset_allocation_col)

            hybrid_set_ppt_cell_text(
                target_cell,
                paste_text,
                font_name=PPT_FONT_NAME,
                font_size=PPT_FONT_SIZE
            )

            total_updated_count += 1
            log(f"Slide {slide_index}: Updated Asset Allocation for: {scheme_name}")

    prs.save(output_path)

    log("")
    log("PowerPoint file saved successfully.")
    log(f"Output file: {output_path}")
    log(f"Total PPT rows updated across all slides: {total_updated_count}")

    if not_found:
        log("")
        log("Schemes not found in Excel:")
        for item in not_found:
            log(f" - {item}")

    return total_updated_count


# ============================================================
# WORKER THREAD
# ============================================================

class hybrid_ProcessWorker(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(int, str)
    error_signal = pyqtSignal(str)

    def __init__(self, excel_path, ppt_path, output_path):
        super().__init__()
        self.excel_path = excel_path
        self.ppt_path = ppt_path
        self.output_path = output_path

    def log(self, message):
        self.log_signal.emit(message)

    def run(self):
        try:
            self.log("==============================================")
            self.log("Process started")
            self.log("==============================================")
            self.log(f"Excel File: {self.excel_path}")
            self.log(f"PPT File: {self.ppt_path}")
            self.log(f"Output File: {self.output_path}")
            self.log("")

            excel_data = hybrid_read_excel_asset_allocations(self.excel_path, self.log)

            if not excel_data:
                raise Exception("No fund asset allocation data found in Excel.")

            self.log("")

            updated_count = hybrid_update_ppt_from_excel(
                ppt_path=self.ppt_path,
                excel_data=excel_data,
                output_path=self.output_path,
                log=self.log
            )

            self.log("")
            self.log("==============================================")
            self.log("Process completed successfully")
            self.log("==============================================")

            self.finished_signal.emit(updated_count, self.output_path)

        except Exception as error:
            error_message = str(error)
            details = traceback.format_exc()

            self.log("")
            self.log("ERROR:")
            self.log(error_message)
            self.log("")
            self.log(details)

            self.error_signal.emit(error_message)


# ============================================================
# PYQT6 GUI
# ============================================================

class hybrid_ExcelToPPTGui(QWidget):
    def __init__(self):
        super().__init__()

        self.worker = None

        self.setWindowTitle("Excel to PPT Asset Allocation Updater")
        self.resize(1200, 650)

        self.excel_input = QLineEdit()
        self.ppt_input = QLineEdit()
        self.output_input = QLineEdit()

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        self.start_button = QPushButton("Start Process")
        self.clear_log_button = QPushButton("Clear Log")

        self.build_ui()

    def build_ui(self):
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(12)

        left_panel = self.build_log_panel()
        right_panel = self.build_process_panel()

        main_layout.addWidget(left_panel, stretch=3)
        main_layout.addWidget(right_panel, stretch=2)

        self.setLayout(main_layout)

    def build_log_panel(self):
        group = QGroupBox("Log")
        layout = QVBoxLayout()

        self.log_box.setStyleSheet("""
            QTextEdit {
                background-color: #111111;
                color: #00ff99;
                font-family: Consolas;
                font-size: 11pt;
                border: 1px solid #444444;
            }
        """)

        layout.addWidget(self.log_box)
        group.setLayout(layout)

        return group

    def build_process_panel(self):
        group = QGroupBox("Process")
        group.setMinimumWidth(430)

        layout = QVBoxLayout()
        layout.setSpacing(14)

        title = QLabel("Excel to PPT Update Process")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #003B5C;
                padding: 8px;
            }
        """)

        layout.addWidget(title)

        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(separator)

        layout.addWidget(self.build_file_picker_row(
            label_text="Excel File",
            line_edit=self.excel_input,
            button_text="Browse",
            button_handler=self.browse_excel
        ))

        layout.addWidget(self.build_file_picker_row(
            label_text="PowerPoint File",
            line_edit=self.ppt_input,
            button_text="Browse",
            button_handler=self.browse_ppt
        ))

        layout.addWidget(self.build_file_picker_row(
            label_text="Output PPT File",
            line_edit=self.output_input,
            button_text="Save As",
            button_handler=self.browse_output
        ))

        self.start_button.setMinimumHeight(42)
        self.start_button.clicked.connect(self.start_process)
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #0B5CAD;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #084A8A;
            }
            QPushButton:disabled {
                background-color: #999999;
            }
        """)

        self.clear_log_button.setMinimumHeight(34)
        self.clear_log_button.clicked.connect(self.clear_log)

        layout.addWidget(self.start_button)
        layout.addWidget(self.clear_log_button)

        info = QLabel(
            "Process Details:\n"
            "1. Searches every slide for table object SnapshotTab.\n"
            "2. Matches PPT Scheme Name with Excel Fund Name.\n"
            "3. Ignores symbols and special characters while matching.\n"
            "4. Reads Asset Class from Excel column D.\n"
            "5. Reads % value from Excel column E.\n"
            "6. Reads formula calculated values using Excel COM on Windows.\n"
            "7. Pastes Asset Class + Value in Asset Allocation(%).\n"
            "8. Applies Ubuntu font and size 6."
        )
        info.setWordWrap(True)
        info.setStyleSheet("""
            QLabel {
                background-color: #F4F8FB;
                border: 1px solid #D4E2EF;
                border-radius: 6px;
                padding: 10px;
                color: #222222;
                font-size: 12px;
            }
        """)

        layout.addWidget(info)
        layout.addStretch()

        group.setLayout(layout)

        return group

    def build_file_picker_row(self, label_text, line_edit, button_text, button_handler):
        container = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel(label_text)
        label.setStyleSheet("font-weight: bold;")

        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 0, 0, 0)

        line_edit.setPlaceholderText(f"Select {label_text}")
        line_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        button = QPushButton(button_text)
        button.clicked.connect(button_handler)
        button.setMinimumWidth(90)

        row_layout.addWidget(line_edit)
        row_layout.addWidget(button)

        layout.addWidget(label)
        layout.addLayout(row_layout)

        container.setLayout(layout)

        return container

    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm *.xlsb *.xls);;All Files (*.*)"
        )

        if file_path:
            self.excel_input.setText(file_path)

    def browse_ppt(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select PowerPoint File",
            "",
            "PowerPoint Files (*.pptx);;All Files (*.*)"
        )

        if file_path:
            self.ppt_input.setText(file_path)
            self.auto_set_output_path(file_path)

    def browse_output(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output PowerPoint File",
            "",
            "PowerPoint Files (*.pptx);;All Files (*.*)"
        )

        if file_path:
            if not file_path.lower().endswith(".pptx"):
                file_path += ".pptx"

            self.output_input.setText(file_path)

    def auto_set_output_path(self, ppt_path):
        if not ppt_path:
            return

        folder = os.path.dirname(ppt_path)
        file_name = os.path.basename(ppt_path)
        name, ext = os.path.splitext(file_name)

        output_path = os.path.join(folder, f"{name}_updated{ext}")

        self.output_input.setText(output_path)

    def validate_inputs(self):
        excel_path = self.excel_input.text().strip()
        ppt_path = self.ppt_input.text().strip()
        output_path = self.output_input.text().strip()

        if not excel_path:
            QMessageBox.critical(self, "Missing Excel File", "Please select an Excel file.")
            return False

        if not os.path.exists(excel_path):
            QMessageBox.critical(self, "Invalid Excel File", "Selected Excel file does not exist.")
            return False

        if not ppt_path:
            QMessageBox.critical(self, "Missing PPT File", "Please select a PowerPoint file.")
            return False

        if not os.path.exists(ppt_path):
            QMessageBox.critical(self, "Invalid PPT File", "Selected PPT file does not exist.")
            return False

        if not output_path:
            QMessageBox.critical(self, "Missing Output File", "Please select output PPT file path.")
            return False

        if not output_path.lower().endswith(".pptx"):
            QMessageBox.critical(self, "Invalid Output File", "Output file must be a .pptx file.")
            return False

        return True

    def start_process(self):
        if not self.validate_inputs():
            return

        self.clear_log()
        self.start_button.setEnabled(False)

        excel_path = self.excel_input.text().strip()
        ppt_path = self.ppt_input.text().strip()
        output_path = self.output_input.text().strip()

        self.worker = hybrid_ProcessWorker(
            excel_path=excel_path,
            ppt_path=ppt_path,
            output_path=output_path
        )

        self.worker.log_signal.connect(self.append_log)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.error_signal.connect(self.process_failed)

        self.worker.start()

    def append_log(self, message):
        self.log_box.append(message)

    def clear_log(self):
        self.log_box.clear()

    def process_finished(self, updated_count, output_path):
        self.start_button.setEnabled(True)

        QMessageBox.information(
            self,
            "Success",
            f"Process completed successfully.\n\n"
            f"Rows updated across all slides: {updated_count}\n\n"
            f"Saved file:\n{output_path}"
        )

    def process_failed(self, error_message):
        self.start_button.setEnabled(True)

        QMessageBox.critical(
            self,
            "Error",
            error_message
        )


# ============================================================
# MAIN
# ============================================================

def hybrid_main():
    app = QApplication(sys.argv)

    window = hybrid_ExcelToPPTGui()
    window.show()

    sys.exit(app.exec())


if False and __name__ == "__main__":
    hybrid_main()



# ============================================================
# SECTION 1 LOGIC: Snapshot Market Cap Updater
# ============================================================
MAIN_TABLE_DEFAULT = "SnapshotTab"
INV_STYLE_PREFIX_DEFAULT = "InvStyle"
CAP_ORDER = ["large", "mid", "small"]
CAP_LABELS = {"large": ["large cap", "largecap", "large"], "mid": ["mid cap", "midcap", "mid"], "small": ["small cap", "smallcap", "small"]}

def normalize_text(value):
    if value is None: return ""
    text = str(value).replace("&amp;amp;", "&amp;").lower().replace("&amp;", "and")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def normalize_for_match(value):
    remove_words = {"bandhan", "mutual", "fund", "regular", "direct", "growth", "plan", "option", "idcw", "dividend", "scheme", "the", "and"}
    return " ".join(w for w in normalize_text(value).split() if w not in remove_words)

def clean_scheme_name(value):
    text = str(value or "").replace("&amp;amp;", "&amp;")
    text = re.sub(r"[£¥§]+", "", text)
    return re.sub(r"\s+", " ", text).strip()

def similarity(a, b):
    a = normalize_for_match(a); b = normalize_for_match(b)
    return SequenceMatcher(None, a, b).ratio() if a and b else 0.0

def round_percent(value):
    if value is None: return None
    has_percent_symbol = False
    if isinstance(value, str):
        raw = value.strip()
        if not raw: return None
        has_percent_symbol = "%" in raw
        raw = raw.replace("%", "").replace(",", "").strip()
        match = re.search(r"-?\d+(\.\d+)?", raw)
        if not match: return None
        number = Decimal(match.group(0))
    else:
        number = Decimal(str(value))
    if not has_percent_symbol and number <= Decimal("1"):
        number *= Decimal("100")
    rounded = number.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return f"{int(rounded)}%"

def parse_number(value):
    if value is None: return None
    if isinstance(value, (int, float, Decimal)): return value
    text = str(value).strip().replace(",", "")
    if not text: return None
    match = re.search(r"-?\d+(\.\d+)?\s*%?", text)
    return match.group(0) if match else None

def is_cap_label(value, cap_type):
    text = normalize_text(value)
    return bool(text) and any(normalize_text(label) in text for label in CAP_LABELS[cap_type])

def get_shape_name(shape):
    try: return shape.name or ""
    except Exception: return ""

def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in iter_shapes(shape.shapes): yield child

def find_shape_by_name(slide, shape_name):
    target = normalize_text(shape_name)
    for shape in iter_shapes(slide.shapes):
        if normalize_text(get_shape_name(shape)) == target: return shape
    return None

def find_table_shape_by_name(slide, table_shape_name):
    target = normalize_text(table_shape_name)
    for shape in iter_shapes(slide.shapes):
        if getattr(shape, "has_table", False) and normalize_text(get_shape_name(shape)) == target: return shape
    return None

def get_cell_text(cell):
    try: return cell.text.strip() if cell.text else ""
    except Exception: return ""

def set_cell_text_keep_format(cell, new_text):
    try:
        text_frame = cell.text_frame
        if not text_frame.paragraphs:
            cell.text = new_text; return True
        paragraph = text_frame.paragraphs[0]
        if not paragraph.runs: paragraph.text = new_text
        else:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]: run.text = ""
        return True
    except Exception:
        try: cell.text = new_text; return True
        except Exception: return False

def table_cell_text(table, row_idx, col_idx):
    try: return get_cell_text(table.cell(row_idx, col_idx))
    except Exception: return ""

def find_scheme_name_column(table, log=None):
    row_count = len(table.rows); col_count = len(table.columns); scan_rows = min(row_count, 12)
    for r in range(scan_rows):
        for c in range(col_count):
            text = normalize_text(table_cell_text(table, r, c))
            if text and ("scheme name" in text or text in {"scheme", "scheme names", "fund name", "fund names"}):
                if log: log(f"Detected Scheme Name column: row={r}, col={c}")
                return c, r
    for r in range(scan_rows):
        for c in range(col_count):
            text = normalize_text(table_cell_text(table, r, c))
            if text == "category" or "category" in text:
                if c + 1 < col_count:
                    if log: log(f"Using column after Category as Scheme Name: col={c + 1}")
                    return c + 1, r
    best_col = None; best_score = 0; fund_words = ["fund", "scheme", "bandhan", "cap", "equity", "debt", "hybrid"]
    for c in range(col_count):
        score = 0
        for r in range(scan_rows, row_count):
            text = normalize_text(table_cell_text(table, r, c))
            if text and any(word in text for word in fund_words): score += 1
        if score > best_score: best_score = score; best_col = c
    if best_col is not None and best_score > 0:
        if log: log(f"Detected Scheme Name by fund-like values: col={best_col}, score={best_score}")
        return best_col, 0
    return None, None

def is_header_or_invalid_scheme_name(value):
    text = normalize_text(value)
    invalid = {"scheme name", "scheme", "fund name", "fund names", "category", "about the fund", "investment style", "benchmark", "top 10 holdings", "top 10 industry", "ratios"}
    return (not text) or text in invalid or "equity snapshot" in text or "category scheme name" in text

def read_scheme_rows_from_snapshot_table(table, log=None):
    scheme_col, header_row = find_scheme_name_column(table, log)
    if scheme_col is None: raise RuntimeError("Scheme Name column was not found in SnapshotTab.")
    result = []; inv_no = 1; start_row = header_row + 1 if header_row is not None else 0
    for r in range(start_row, len(table.rows)):
        scheme_name = clean_scheme_name(table_cell_text(table, r, scheme_col))
        if is_header_or_invalid_scheme_name(scheme_name): continue
        result.append({"ppt_row_index": r, "inv_style_number": inv_no, "scheme_name": scheme_name}); inv_no += 1
    if log:
        log(f"Scheme Name column index: {scheme_col}"); log(f"Scheme rows detected: {len(result)}")
    return result

def collect_excel_files(excel_folder):
    files = []
    for root, _, filenames in os.walk(excel_folder):
        for name in filenames:
            lower = name.lower()
            if lower.startswith("~$"): continue
            if lower.endswith((".xlsx", ".xlsm")): files.append(Path(root) / name)
    return files

def find_best_excel_file(scheme_name, excel_files):
    best_file = None; best_score = 0.0
    for excel_file in excel_files:
        score = similarity(scheme_name, excel_file.stem)
        if score > best_score: best_score = score; best_file = excel_file
    return (best_file, best_score) if best_score >= 0.30 else (None, best_score)

def extract_market_caps_from_excel(excel_path, sheet_name=None):
    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if sheet_name:
            if sheet_name not in workbook.sheetnames: raise RuntimeError(f"Sheet '{sheet_name}' not found in {excel_path.name}.")
            sheets = [workbook[sheet_name]]
        else:
            sheets = workbook.worksheets
        found = {"large": None, "mid": None, "small": None}
        for ws in sheets:
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, 500), min_col=1, max_col=min(ws.max_column or 0, 100)):
                values = [cell.value for cell in row]
                for idx, value in enumerate(values):
                    for cap_type in CAP_ORDER:
                        if found[cap_type] is not None or not is_cap_label(value, cap_type): continue
                        number = None
                        for right_value in values[idx + 1:]:
                            number = parse_number(right_value)
                            if number is not None: break
                        if number is None:
                            for left_value in reversed(values[:idx]):
                                number = parse_number(left_value)
                                if number is not None: break
                        rounded = round_percent(number)
                        if rounded: found[cap_type] = rounded
                if all(found.values()): return found
        return found
    finally:
        workbook.close()

def has_text(shape): return getattr(shape, "has_text_frame", False)
def get_shape_text(shape): return shape.text_frame.text.strip() if has_text(shape) else ""

def set_shape_text_keep_format(shape, new_text):
    if not has_text(shape): return False
    tf = shape.text_frame
    if not tf.paragraphs: tf.text = new_text; return True
    p = tf.paragraphs[0]
    if not p.runs: p.text = new_text
    else:
        p.runs[0].text = new_text
        for run in p.runs[1:]: run.text = ""
    return True

def collect_text_shapes(shape):
    result = []
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in iter_shapes(shape.shapes):
            if has_text(child): result.append(child)
    elif has_text(shape): result.append(shape)
    return result

def collect_table_shapes(shape):
    result = []
    if getattr(shape, "has_table", False): result.append(shape)
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in iter_shapes(shape.shapes):
            if getattr(child, "has_table", False): result.append(child)
    return result

def update_inv_style_table(table_shape, market_caps):
    table = table_shape.table; row_count = len(table.rows); col_count = len(table.columns)
    label_rows = {}; percent_cells = []
    for r in range(row_count):
        for c in range(col_count):
            text = table_cell_text(table, r, c)
            if re.search(r"\d+\s*%", text): percent_cells.append((r, c))
            for cap_type in CAP_ORDER:
                if cap_type not in label_rows and is_cap_label(text, cap_type): label_rows[cap_type] = r
    updated = 0
    for cap_type in CAP_ORDER:
        value = market_caps.get(cap_type); row_idx = label_rows.get(cap_type)
        if not value or row_idx is None: continue
        target_col = None
        for c in range(col_count):
            if re.search(r"\d+\s*%", table_cell_text(table, row_idx, c)): target_col = c; break
        if target_col is None:
            for c in range(col_count):
                text = table_cell_text(table, row_idx, c)
                if not any(is_cap_label(text, label_type) for label_type in CAP_ORDER): target_col = c; break
        if target_col is not None and set_cell_text_keep_format(table.cell(row_idx, target_col), value): updated += 1
    if updated: return True, f"Updated {updated} values inside InvStyle table using Large/Mid/Small label rows."
    values = [v for v in [market_caps.get("large"), market_caps.get("mid"), market_caps.get("small")] if v]
    for (r, c), value in zip(sorted(percent_cells), values):
        if set_cell_text_keep_format(table.cell(r, c), value): updated += 1
    if updated: return True, f"Updated {updated} values inside InvStyle table using top-to-bottom percentage fallback."
    return False, "InvStyle table found, but percentage cells could not be mapped."

def update_inv_style_shape(inv_shape, market_caps):
    for table_shape in collect_table_shapes(inv_shape):
        ok, msg = update_inv_style_table(table_shape, market_caps)
        if ok: return True, msg
    text_shapes = collect_text_shapes(inv_shape)
    percent_shapes = sorted([s for s in text_shapes if re.search(r"\d+\s*%", get_shape_text(s))], key=lambda s: (s.top, s.left))
    values = [v for v in [market_caps.get("large"), market_caps.get("mid"), market_caps.get("small")] if v]
    updated = 0
    for shape, value in zip(percent_shapes, values):
        set_shape_text_keep_format(shape, value); updated += 1
    if updated: return True, f"Updated {updated} values using editable text top-to-bottom fallback."
    return False, "No editable percentage cells/text boxes found inside InvStyle."

def process_ppt(ppt_path, excel_folder, output_path, table_object_name, inv_style_prefix, excel_sheet_name, log, progress):
    excel_files = collect_excel_files(excel_folder)
    if not excel_files: raise RuntimeError("No Excel files found in selected Excel folder.")
    log(f"Excel files found: {len(excel_files)}"); log(f"Opening PPT: {ppt_path}")
    prs = Presentation(str(ppt_path)); work_items = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        table_shape = find_table_shape_by_name(slide, table_object_name)
        if table_shape is None: continue
        log(f"Found table '{table_object_name}' on slide {slide_idx}")
        for row in read_scheme_rows_from_snapshot_table(table_shape.table, log):
            row["slide_idx"] = slide_idx; row["slide"] = slide; work_items.append(row)
    if not work_items: raise RuntimeError(f"No rows found. Check table object name '{table_object_name}'.")
    total = len(work_items); updated = 0; skipped = 0; log(f"Funds found in PPT table: {total}")
    for index, item in enumerate(work_items, start=1):
        scheme_name = item["scheme_name"]; slide = item["slide"]; inv_style_name = f"{inv_style_prefix}{item['inv_style_number']}"
        log(""); log(f"[{index}/{total}] Slide {item['slide_idx']}"); log(f"Scheme: {scheme_name}"); log(f"InvStyle object: {inv_style_name}")
        inv_shape = find_shape_by_name(slide, inv_style_name)
        if inv_shape is None:
            log(f"SKIP: Shape '{inv_style_name}' not found."); skipped += 1; progress(index, total); continue
        excel_file, score = find_best_excel_file(scheme_name, excel_files)
        if excel_file is None:
            log("SKIP: Matching Excel file not found."); skipped += 1; progress(index, total); continue
        log(f"Matched Excel: {excel_file.name} | score: {score:.2f}")
        market_caps = extract_market_caps_from_excel(excel_file, excel_sheet_name)
        log(f"Excel Market Cap: Large={market_caps.get('large')}, Mid={market_caps.get('mid')}, Small={market_caps.get('small')}")
        if not any(market_caps.values()):
            log("SKIP: Large/Mid/Small Cap values not found in Excel."); skipped += 1; progress(index, total); continue
        ok, msg = update_inv_style_shape(inv_shape, market_caps)
        if ok: log(f"OK: {msg}"); updated += 1
        else: log(f"SKIP: {msg}"); skipped += 1
        progress(index, total)
    if updated == 0: raise RuntimeError("No InvStyle object was updated. Output PPT was not saved.")
    prs.save(str(output_path)); log(""); log("Completed."); log(f"Updated funds: {updated}"); log(f"Skipped funds: {skipped}"); log(f"Saved output PPT: {output_path}")

class MarketCapWorker(QThread):
    log_signal = pyqtSignal(str); progress_signal = pyqtSignal(int); status_signal = pyqtSignal(str); finished_signal = pyqtSignal(bool, str)
    def __init__(self, ppt_path, excel_folder, output_path, table_object_name, inv_style_prefix, excel_sheet_name):
        super().__init__(); self.ppt_path = Path(ppt_path); self.excel_folder = Path(excel_folder); self.output_path = Path(output_path)
        self.table_object_name = table_object_name; self.inv_style_prefix = inv_style_prefix; self.excel_sheet_name = excel_sheet_name.strip() or None
    def run(self):
        try:
            self.status_signal.emit("Running..."); self.progress_signal.emit(0)
            process_ppt(self.ppt_path, self.excel_folder, self.output_path, self.table_object_name, self.inv_style_prefix, self.excel_sheet_name, self.log_signal.emit, lambda c,t: self.progress_signal.emit(0 if t <= 0 else int((c / t) * 100)))
            self.progress_signal.emit(100); self.status_signal.emit("Completed"); self.finished_signal.emit(True, "Process completed successfully.")
        except Exception:
            error = traceback.format_exc(); self.log_signal.emit("ERROR:"); self.log_signal.emit(error); self.status_signal.emit("Failed"); self.finished_signal.emit(False, error)

class MarketCapSection(QWidget):
    def __init__(self, state: Any | None = None):
        super().__init__(); self.state = state; self.worker = None; self.default_output_folder = self._resolve_default_output_folder()
        main_layout = QHBoxLayout(self); main_layout.addWidget(self.create_left_log_panel()); main_layout.addWidget(self.create_right_panel()); self._apply_default_output_path()
    def _app_run_folder(self) -> Path:
        if getattr(sys, "frozen", False): return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent
    def _resolve_default_output_folder(self) -> Path:
        """Return Tab 4 Section 1 output folder.

        In the main app this tab receives a scoped state whose output_root is:
            Bandhan_Output/04_snapshot_automation

        Older builds reused the old Tab 5 folder name 05_marketcap_updated, which
        created an extra folder inside Tab 4. The correct Section 1 folder is:
            01_marketcap_updated
        """
        state_folder = None
        if self.state is not None:
            state_folder = (
                getattr(self.state, "tab4_section1_marketcap_folder", None)
                or getattr(self.state, "tab5_marketcap_folder", None)
                or getattr(self.state, "tab5_updated_folder", None)
            )
        if state_folder:
            folder = Path(state_folder)
        else:
            output_root = getattr(self.state, "output_root", None) if self.state is not None else None
            folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "01_marketcap_updated"
        folder.mkdir(parents=True, exist_ok=True)
        if self.state is not None:
            setattr(self.state, "tab4_section1_marketcap_folder", folder)
            setattr(self.state, "tab5_marketcap_folder", folder)
            setattr(self.state, "tab5_updated_folder", folder)
        return folder
    def _default_output_path(self, ppt_path: Path | None = None) -> Path:
        return self.default_output_folder / (f"{ppt_path.stem}_marketcap_updated.pptx" if ppt_path else "snapshot_marketcap_updated.pptx")
    def _apply_default_output_path(self):
        if hasattr(self, "output_input") and not self.output_input.text().strip(): self.output_input.setText(str(self._default_output_path()))
    def create_left_log_panel(self):
        layout = QVBoxLayout(); title = QLabel("Process Logs"); title.setStyleSheet("font-weight: bold; font-size: 13px;"); self.log_box = QTextEdit(); self.log_box.setReadOnly(True)
        layout.addWidget(title); layout.addWidget(self.log_box); widget = QWidget(); widget.setLayout(layout); widget.setMinimumWidth(380); widget.setMaximumWidth(420); return widget
    def create_right_panel(self):
        layout = QVBoxLayout(); title = QLabel("Snapshot Market Cap Updater"); title.setStyleSheet("font-weight: bold; font-size: 18px;")
        desc = QLabel("Updates the Investment Style market-cap table in Snapshot PPT. Reads Scheme Name from SnapshotTab, finds matching Excel files, and updates InvStyle1, InvStyle2, etc."); desc.setWordWrap(True)
        layout.addWidget(title); layout.addWidget(desc); layout.addWidget(self.create_source_group()); layout.addWidget(self.create_input_group()); layout.addWidget(self.create_settings_group()); layout.addWidget(self.create_output_group()); layout.addWidget(self.create_progress_group()); layout.addStretch(); widget = QWidget(); widget.setLayout(layout); return widget
    def create_source_group(self):
        group = QGroupBox("1. Source Folder / Auto Detect"); layout = QHBoxLayout(); self.source_folder_input = QLineEdit(); self.source_folder_input.setReadOnly(True)
        browse = QPushButton("Browse Source Folder..."); browse.clicked.connect(self.browse_source_folder); auto = QPushButton("Auto Detect Files"); auto.clicked.connect(self.auto_detect_files)
        layout.addWidget(QLabel("Source Folder:")); layout.addWidget(self.source_folder_input); layout.addWidget(browse); layout.addWidget(auto); group.setLayout(layout); return group
    def create_input_group(self):
        group = QGroupBox("2. Input Files"); layout = QVBoxLayout()
        excel_row = QHBoxLayout(); self.excel_folder_input = QLineEdit(); excel_button = QPushButton("Browse Excel Folder..."); excel_button.clicked.connect(self.browse_excel_folder)
        excel_row.addWidget(QLabel("Excel Folder:")); excel_row.addWidget(self.excel_folder_input); excel_row.addWidget(excel_button)
        ppt_row = QHBoxLayout(); self.ppt_input = QLineEdit(); ppt_button = QPushButton("Browse Snapshot PPTX..."); ppt_button.clicked.connect(self.browse_ppt)
        ppt_row.addWidget(QLabel("Snapshot PPTX:")); ppt_row.addWidget(self.ppt_input); ppt_row.addWidget(ppt_button)
        layout.addLayout(excel_row); layout.addLayout(ppt_row); group.setLayout(layout); return group
    def create_settings_group(self):
        group = QGroupBox("3. Update Settings"); layout = QVBoxLayout(); self.sheet_name_input = QLineEdit("Report"); self.sheet_name_input.setPlaceholderText("Default: Report")
        row = QHBoxLayout(); row.addWidget(QLabel("Excel Sheet Name:")); row.addWidget(self.sheet_name_input)
        note = QLabel("PPT Main Table Object Name and Investment Style Object Prefix are hidden and run in background as SnapshotTab and InvStyle."); note.setWordWrap(True)
        layout.addLayout(row); layout.addWidget(note); group.setLayout(layout); return group
    def create_output_group(self):
        group = QGroupBox("4. Output / Run"); layout = QVBoxLayout(); output_row = QHBoxLayout(); self.output_input = QLineEdit(); save = QPushButton("Save Output PPT As..."); save.clicked.connect(self.browse_output)
        output_row.addWidget(QLabel("Output PPTX:")); output_row.addWidget(self.output_input); output_row.addWidget(save)
        run_row = QHBoxLayout(); run_row.addStretch(); self.run_button = QPushButton("Run Market Cap Update"); self.run_button.clicked.connect(self.run_process); run_row.addWidget(self.run_button)
        layout.addLayout(output_row); layout.addLayout(run_row); group.setLayout(layout); return group
    def create_progress_group(self):
        group = QGroupBox("Progress"); layout = QVBoxLayout(); self.status_label = QLabel("Idle"); self.progress_bar = QProgressBar(); self.progress_bar.setValue(0); layout.addWidget(self.status_label); layout.addWidget(self.progress_bar); group.setLayout(layout); return group
    def log(self, message): self.log_box.append(str(message)); self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())
    def browse_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Source Folder", str(Path.home()))
        if folder: self.source_folder_input.setText(folder)
    def browse_excel_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Excel Folder", self.source_folder_input.text() or str(Path.home()))
        if folder: self.excel_folder_input.setText(folder)
    def browse_ppt(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Snapshot PPTX", self.source_folder_input.text() or str(Path.home()), "PowerPoint Files (*.pptx)")
        if file_path:
            self.ppt_input.setText(file_path); self.output_input.setText(str(self._default_output_path(Path(file_path))))
    def browse_output(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Output PPT As", self.output_input.text() or str(self._default_output_path()), "PowerPoint Files (*.pptx)")
        if file_path:
            if not file_path.lower().endswith(".pptx"): file_path += ".pptx"
            self.output_input.setText(file_path)
    def auto_detect_files(self):
        source = self.source_folder_input.text().strip()
        if not source: QMessageBox.warning(self, "Missing Source Folder", "Please select source folder first."); return
        source_path = Path(source)
        if not source_path.exists(): QMessageBox.warning(self, "Invalid Folder", "Selected source folder does not exist."); return
        self.excel_folder_input.setText(str(source_path)); ppt_candidates = []
        for root, _, files in os.walk(source_path):
            for name in files:
                lower = name.lower()
                if lower.startswith("~$") or not lower.endswith(".pptx"): continue
                full_path = Path(root) / name; norm = normalize_text(full_path.stem)
                if "equity snapshot" in norm or "equity snashot" in norm: ppt_candidates.insert(0, full_path)
                elif "hybrid snapshot" in norm or "debt snapshot" in norm or "snapshot" in norm or "snashot" in norm: ppt_candidates.append(full_path)
        if ppt_candidates:
            ppt_path = ppt_candidates[0]; self.ppt_input.setText(str(ppt_path)); self.output_input.setText(str(self._default_output_path(ppt_path))); self.log(f"Auto detected PPT: {ppt_path}"); self.log(f"Excel folder set to: {source_path}")
        else: self.log("No Snapshot PPTX found during auto detect.")
    def validate_inputs(self):
        ppt_path = self.ppt_input.text().strip(); excel_folder = self.excel_folder_input.text().strip(); output_path = self.output_input.text().strip()
        if not ppt_path: return False, "Please select Snapshot PPTX."
        if not Path(ppt_path).exists(): return False, "Selected PPTX does not exist."
        if not excel_folder: return False, "Please select Excel folder."
        if not Path(excel_folder).exists(): return False, "Selected Excel folder does not exist."
        if not output_path: return False, "Please select output PPTX path."
        if Path(ppt_path).resolve() == Path(output_path).resolve(): return False, "Output PPTX cannot be same as input PPTX."
        return True, ""
    def set_controls_enabled(self, enabled): self.run_button.setEnabled(enabled)
    def run_process(self):
        valid, msg = self.validate_inputs()
        if not valid: QMessageBox.warning(self, "Validation Error", msg); return
        self.log_box.clear(); self.progress_bar.setValue(0); self.status_label.setText("Starting..."); self.set_controls_enabled(False)
        self.worker = MarketCapWorker(self.ppt_input.text().strip(), self.excel_folder_input.text().strip(), self.output_input.text().strip(), MAIN_TABLE_DEFAULT, INV_STYLE_PREFIX_DEFAULT, self.sheet_name_input.text().strip())
        self.worker.log_signal.connect(self.log); self.worker.progress_signal.connect(self.progress_bar.setValue); self.worker.status_signal.connect(self.status_label.setText); self.worker.finished_signal.connect(self.on_finished); self.worker.start()
    def on_finished(self, success, message):
        self.set_controls_enabled(True)
        if success: QMessageBox.information(self, "Done", message)
        else: QMessageBox.critical(self, "Failed", "Process failed. Please check logs.")


# ============================================================
# SECTION 2 GUI: Asset Allocation Excel Generator
# ============================================================
class AssetAllocationExcelSection(QWidget):
    def __init__(self, state: Any | None = None):
        super().__init__(); self.state = state; self.worker = None
        main_layout = QHBoxLayout(self); main_layout.addWidget(self.create_left_log_panel()); main_layout.addWidget(self.create_right_panel()); self._apply_default_output_path()
    def _app_run_folder(self) -> Path:
        return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    def _default_output_path(self) -> Path:
        output_root = getattr(self.state, "output_root", None) if self.state is not None else None
        folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "02_asset_allocation_excel"; folder.mkdir(parents=True, exist_ok=True)
        return folder / "asset_allocation_output.xlsx"
    def _apply_default_output_path(self):
        if hasattr(self, "output_input") and not self.output_input.text().strip(): self.output_input.setText(str(self._default_output_path()))
    def create_left_log_panel(self):
        layout = QVBoxLayout(); title = QLabel("Process Logs"); title.setStyleSheet("font-weight: bold; font-size: 13px;"); self.log_box = QTextEdit(); self.log_box.setReadOnly(True)
        layout.addWidget(title); layout.addWidget(self.log_box); widget = QWidget(); widget.setLayout(layout); widget.setMinimumWidth(380); widget.setMaximumWidth(420); return widget
    def create_right_panel(self):
        layout = QVBoxLayout(); title = QLabel("Asset Allocation Excel Generator"); title.setStyleSheet("font-weight: bold; font-size: 18px;")
        desc = QLabel("Section 2: Generates the Asset Allocation Excel file from Consolidate and Mapping files."); desc.setWordWrap(True)
        layout.addWidget(title); layout.addWidget(desc); layout.addWidget(self.create_input_group()); layout.addWidget(self.create_fund_group()); layout.addWidget(self.create_output_group()); layout.addWidget(self.create_progress_group()); layout.addStretch(); widget = QWidget(); widget.setLayout(layout); return widget
    def create_input_group(self):
        group = QGroupBox("1. Input Files"); layout = QVBoxLayout()
        r1 = QHBoxLayout(); self.consolidate_input = QLineEdit(); b1 = QPushButton("Browse Consolidate File..."); b1.clicked.connect(self.browse_consolidate); r1.addWidget(QLabel("Consolidate File:")); r1.addWidget(self.consolidate_input); r1.addWidget(b1)
        r2 = QHBoxLayout(); self.mapping_input = QLineEdit(); b2 = QPushButton("Browse Mapping File..."); b2.clicked.connect(self.browse_mapping); r2.addWidget(QLabel("Mapping File:")); r2.addWidget(self.mapping_input); r2.addWidget(b2)
        layout.addLayout(r1); layout.addLayout(r2); group.setLayout(layout); return group
    def create_fund_group(self):
        group = QGroupBox("2. Fund Names")
        layout = QVBoxLayout()

        self.fund_input = QTextEdit()

        default_fund_names = (
            "Bandhan Arbitrage Fund\n"
            "Bandhan Equity Savings Fund\n"
            "Bandhan Conservative Hybrid Fund\n"
            "Bandhan Balanced Advantage Fund\n"
            "Bandhan Aggressive Hybrid Fund\n"
            "Bandhan Retirement Fund\n"
            "Bandhan Multi Asset Allocation Fund"
        )

        self.fund_input.setPlainText(default_fund_names)

        self.fund_input.setPlaceholderText(
            "Enter one fund name per line.\n"
            "Example:\n"
            "Bandhan Arbitrage Fund\n"
            "Bandhan Equity Savings Fund\n"
            "Bandhan Conservative Hybrid Fund\n"
            "Bandhan Multi Asset Allocation Fund"
        )

        layout.addWidget(self.fund_input)
        group.setLayout(layout)
        return group
    def create_output_group(self):
        group = QGroupBox("3. Output / Run"); layout = QVBoxLayout(); r = QHBoxLayout(); self.output_input = QLineEdit(); b = QPushButton("Save Output Excel As..."); b.clicked.connect(self.browse_output); r.addWidget(QLabel("Output Excel:")); r.addWidget(self.output_input); r.addWidget(b)
        rr = QHBoxLayout(); rr.addStretch(); self.process_button = QPushButton("Generate Asset Allocation Excel"); self.process_button.clicked.connect(self.start_process); rr.addWidget(self.process_button); layout.addLayout(r); layout.addLayout(rr); group.setLayout(layout); return group
    def create_progress_group(self):
        group = QGroupBox("Progress"); layout = QVBoxLayout(); self.status_label = QLabel("Idle"); self.progress_bar = QProgressBar(); self.progress_bar.setValue(0); layout.addWidget(self.status_label); layout.addWidget(self.progress_bar); group.setLayout(layout); return group
    def log(self, message): self.log_box.append(str(message)); self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())
    def browse_consolidate(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Consolidate Excel File", "", "Excel Files (*.xlsx *.xlsm *.xls)")
        if file_path: self.consolidate_input.setText(file_path)
    def browse_mapping(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Mapping Excel File", "", "Excel Files (*.xlsx *.xlsm *.xls)")
        if file_path: self.mapping_input.setText(file_path)
    def browse_output(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Output Excel File", self.output_input.text() or str(self._default_output_path()), "Excel Files (*.xlsx)")
        if file_path:
            if not file_path.lower().endswith(".xlsx"): file_path += ".xlsx"
            self.output_input.setText(file_path)
    def start_process(self):
        consolidate_path = self.consolidate_input.text().strip(); mapping_path = self.mapping_input.text().strip(); output_path = self.output_input.text().strip(); fund_names = asset_parse_fund_names(self.fund_input.toPlainText().strip())
        if not consolidate_path: QMessageBox.warning(self, "Missing File", "Please select Consolidate file."); return
        if not mapping_path: QMessageBox.warning(self, "Missing File", "Please select Mapping file."); return
        if not output_path: QMessageBox.warning(self, "Missing Output", "Please select output file path."); return
        if not fund_names: QMessageBox.warning(self, "Missing Fund Names", "Please enter at least one fund name."); return
        self.log_box.clear(); self.progress_bar.setValue(0); self.status_label.setText("Running..."); self.log("Starting process..."); self.log(f"Total funds entered: {len(fund_names)}"); self.process_button.setEnabled(False)
        self.worker = asset_ExcelWorker(consolidate_path, mapping_path, output_path, fund_names); self.worker.log.connect(self.log); self.worker.success.connect(self.on_success); self.worker.error.connect(self.on_error); self.worker.start()
    def on_success(self, output_path, warning_text):
        self.process_button.setEnabled(True); self.progress_bar.setValue(100); self.status_label.setText("Completed"); self.log("Process completed."); self.log(warning_text)
        QMessageBox.information(self, "Success", f"Output file created successfully:\n{output_path}")
    def on_error(self, message):
        self.process_button.setEnabled(True); self.status_label.setText("Failed"); self.log("ERROR:"); self.log(message); QMessageBox.critical(self, "Error", message)

# ============================================================
# SECTION 3 GUI: Hybrid Asset Allocation Snapshot Update
# ============================================================
class HybridAssetSnapshotSection(QWidget):
    def __init__(self, state: Any | None = None):
        super().__init__(); self.state = state; self.worker = None
        main_layout = QHBoxLayout(self); main_layout.addWidget(self.create_left_log_panel()); main_layout.addWidget(self.create_right_panel()); self._apply_default_output_path()
    def _app_run_folder(self) -> Path:
        return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    def _default_output_path(self, ppt_path: Path | None = None) -> Path:
        output_root = getattr(self.state, "output_root", None) if self.state is not None else None
        folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "03_hybrid_asset_snapshot_updated"; folder.mkdir(parents=True, exist_ok=True)
        return folder / (f"{ppt_path.stem}_asset_allocation_updated.pptx" if ppt_path else "hybrid_asset_allocation_updated.pptx")
    def _apply_default_output_path(self):
        if hasattr(self, "output_input") and not self.output_input.text().strip(): self.output_input.setText(str(self._default_output_path()))
    def create_left_log_panel(self):
        layout = QVBoxLayout(); title = QLabel("Process Logs"); title.setStyleSheet("font-weight: bold; font-size: 13px;"); self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(title); layout.addWidget(self.log_box); widget = QWidget(); widget.setLayout(layout); widget.setMinimumWidth(380); widget.setMaximumWidth(420); return widget
    def create_right_panel(self):
        layout = QVBoxLayout(); title = QLabel("Hybrid Asset Allocation Snapshot Update"); title.setStyleSheet("font-weight: bold; font-size: 18px;"); desc = QLabel("Section 3: Inserts Asset Allocation data into the Hybrid Snapshot PPT."); desc.setWordWrap(True); layout.addWidget(title); layout.addWidget(desc); layout.addWidget(self.create_input_group()); layout.addWidget(self.create_output_group()); layout.addWidget(self.create_progress_group()); layout.addStretch(); widget = QWidget(); widget.setLayout(layout); return widget
    def create_input_group(self):
        group = QGroupBox("1. Input Files"); layout = QVBoxLayout(); r1 = QHBoxLayout(); self.excel_input = QLineEdit(); b1 = QPushButton("Browse Asset Allocation Excel..."); b1.clicked.connect(self.browse_excel); r1.addWidget(QLabel("Excel File:")); r1.addWidget(self.excel_input); r1.addWidget(b1); r2 = QHBoxLayout(); self.ppt_input = QLineEdit(); b2 = QPushButton("Browse Hybrid Snapshot PPTX..."); b2.clicked.connect(self.browse_ppt); r2.addWidget(QLabel("Snapshot PPTX:")); r2.addWidget(self.ppt_input); r2.addWidget(b2); layout.addLayout(r1); layout.addLayout(r2); group.setLayout(layout); return group
    def create_output_group(self):
        group = QGroupBox("2. Output / Run"); layout = QVBoxLayout(); r = QHBoxLayout(); self.output_input = QLineEdit(); b = QPushButton("Save Output PPT As..."); b.clicked.connect(self.browse_output); r.addWidget(QLabel("Output PPTX:")); r.addWidget(self.output_input); r.addWidget(b); rr = QHBoxLayout(); rr.addStretch(); self.start_button = QPushButton("Run Hybrid Asset Snapshot Update"); self.start_button.clicked.connect(self.start_process); rr.addWidget(self.start_button); layout.addLayout(r); layout.addLayout(rr); group.setLayout(layout); return group
    def create_progress_group(self):
        group = QGroupBox("Progress"); layout = QVBoxLayout(); self.status_label = QLabel("Idle"); self.progress_bar = QProgressBar(); self.progress_bar.setValue(0); layout.addWidget(self.status_label); layout.addWidget(self.progress_bar); group.setLayout(layout); return group
    def append_log(self, message): self.log_box.append(str(message)); self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())
    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xlsm *.xlsb *.xls);;All Files (*.*)")
        if file_path: self.excel_input.setText(file_path)
    def browse_ppt(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PowerPoint File", "", "PowerPoint Files (*.pptx);;All Files (*.*)")
        if file_path: self.ppt_input.setText(file_path); self.output_input.setText(str(self._default_output_path(Path(file_path))))
    def browse_output(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Output PowerPoint File", self.output_input.text() or str(self._default_output_path()), "PowerPoint Files (*.pptx);;All Files (*.*)")
        if file_path:
            if not file_path.lower().endswith(".pptx"): file_path += ".pptx"
            self.output_input.setText(file_path)
    def validate_inputs(self):
        excel_path = self.excel_input.text().strip(); ppt_path = self.ppt_input.text().strip(); output_path = self.output_input.text().strip()
        if not excel_path: QMessageBox.critical(self, "Missing Excel File", "Please select an Excel file."); return False
        if not os.path.exists(excel_path): QMessageBox.critical(self, "Invalid Excel File", "Selected Excel file does not exist."); return False
        if not ppt_path: QMessageBox.critical(self, "Missing PPT File", "Please select a PowerPoint file."); return False
        if not os.path.exists(ppt_path): QMessageBox.critical(self, "Invalid PPT File", "Selected PPT file does not exist."); return False
        if not output_path: QMessageBox.critical(self, "Missing Output File", "Please select output PPT file path."); return False
        if not output_path.lower().endswith(".pptx"): QMessageBox.critical(self, "Invalid Output File", "Output file must be a .pptx file."); return False
        return True
    def start_process(self):
        if not self.validate_inputs(): return
        self.log_box.clear(); self.progress_bar.setValue(0); self.status_label.setText("Running..."); self.start_button.setEnabled(False)
        self.worker = hybrid_ProcessWorker(excel_path=self.excel_input.text().strip(), ppt_path=self.ppt_input.text().strip(), output_path=self.output_input.text().strip())
        self.worker.log_signal.connect(self.append_log); self.worker.finished_signal.connect(self.process_finished); self.worker.error_signal.connect(self.process_failed); self.worker.start()
    def process_finished(self, updated_count, output_path):
        self.start_button.setEnabled(True); self.progress_bar.setValue(100); self.status_label.setText("Completed"); QMessageBox.information(self, "Process Completed", f"Updated rows: {updated_count}\n\nSaved file:\n{output_path}")
    def process_failed(self, error_message):
        self.start_button.setEnabled(True); self.status_label.setText("Failed"); QMessageBox.critical(self, "Process Failed", error_message)

class CombinedMarketAssetHybridWidget(QWidget):
    def __init__(self, state: Any | None = None):
        super().__init__(); self.state = state; layout = QVBoxLayout(self); title = QLabel("Snapshot Automation"); title.setStyleSheet("font-weight: bold; font-size: 18px;"); tabs = QTabWidget(); tabs.addTab(MarketCapSection(state), "Section 1 - Market Cap"); tabs.addTab(AssetAllocationExcelSection(state), "Section 2 - Asset Allocation Excel"); tabs.addTab(HybridAssetSnapshotSection(state), "Section 3 - Hybrid Asset Snapshot"); layout.addWidget(title); layout.addWidget(tabs)

Tab4Widget = CombinedMarketAssetHybridWidget
Tab5Widget = CombinedMarketAssetHybridWidget
MarketCapWidget = CombinedMarketAssetHybridWidget
SnapshotMarketCapWidget = CombinedMarketAssetHybridWidget
SnapshotMarketCapUpdater = CombinedMarketAssetHybridWidget

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("Bandhan Factsheet Automation - MarketCap / Asset / Hybrid"); self.resize(1280, 720); self.setCentralWidget(CombinedMarketAssetHybridWidget())

def main():
    app = QApplication(sys.argv); window = MainWindow(); window.show(); sys.exit(app.exec())

if __name__ == "__main__":
    main()
