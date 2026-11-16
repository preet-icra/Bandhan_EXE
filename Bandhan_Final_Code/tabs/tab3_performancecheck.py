from __future__ import annotations

import html
import re
import sys
import traceback
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pptx import Presentation

from PyQt6.QtCore import QObject, QRunnable, QThreadPool, pyqtSignal, pyqtSlot
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
    QScrollArea,
    QSplitter,
    QTextEdit,
QTabWidget,
    QVBoxLayout,
    QWidget,
)


APP_TITLE = "Combined SIP Performance + Performance Scheme Checks"

EXCEL_SCHEME_COL_INDEX = 0
EXCEL_LAUNCH_REGULAR_COL_INDEX = 4
EXCEL_LAUNCH_DIRECT_COL_INDEX = 5

SIP_PERIODS = [
    ("year", 1),
    ("year", 3),
    ("year", 5),
    ("year", 7),
    ("year", 10),
]

SCHEME_STOPWORDS = {
    "bandhan",
    "regular",
    "direct",
    "plan",
    "growth",
    "idcw",
    "option",
    "fund",
    "scheme",
    "the",
}

KNOWN_HTML_TAGS = {
    "sup",
    "sub",
    "b",
    "strong",
    "i",
    "em",
    "u",
    "br",
    "p",
    "div",
    "span",
}

FOOTNOTE_SYMBOL_RE = re.compile(r"[£†‡*§¥¢ß@^$#]+")


# =========================================================
# WORKER
# =========================================================
class WorkerSignals(QObject):
    started = pyqtSignal()
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)


class FunctionWorker(QRunnable):
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


@dataclass
class CheckResult:
    output_path: Path
    rows: int
    fix_rows: int
    ok_rows: int
    review_rows: int


@dataclass(frozen=True)
class Period:
    unit: str
    value: int

    @property
    def label(self) -> str:
        if self.unit == "day":
            return f"{self.value} Day" if self.value == 1 else f"{self.value} Days"

        if self.unit == "month":
            return f"{self.value} Month" if self.value == 1 else f"{self.value} Months"

        return f"{self.value} Year" if self.value == 1 else f"{self.value} Years"

    @property
    def sort_key(self) -> tuple[int, int]:
        order = {
            "day": 0,
            "month": 1,
            "year": 2,
        }

        return order.get(self.unit, 99), self.value


@dataclass
class SchemeLaunch:
    scheme_name: str
    launch_date: Optional[date]


@dataclass
class TableBlock:
    scheme_text: str
    slide: int
    shape_name: str
    start_row: int
    end_row: int
    period_columns: dict[Period, list[int]]
    rows_by_period: dict[Period, list[str]]


# =========================================================
# TEXT / DATE / MATCHING HELPERS
# =========================================================
def decode_repeated_html(value) -> str:
    text = "" if value is None else str(value)
    previous = None

    while previous != text:
        previous = text
        text = html.unescape(text)

    return text


def normalize_text(value) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = decode_repeated_html(value)
    text = text.replace("\u00a0", " ")
    text = text.replace("–", "-")
    text = text.replace("—", "-")

    text = re.sub(
        r"<\s*(sup|sub)\b[^>]*>.*?<\s*/\s*\1\s*>",
        " ",
        text,
        flags=re.I | re.S,
    )
    text = re.sub(
        r"&lt;\s*(sup|sub)\b[^&]*&gt;.*?&lt;\s*/\s*\1\s*&gt;",
        " ",
        text,
        flags=re.I | re.S,
    )

    for tag in KNOWN_HTML_TAGS:
        text = re.sub(rf"</?\s*{tag}\b[^>]*>", " ", text, flags=re.I)
        text = re.sub(rf"&lt;/?\s*{tag}\b[^&]*&gt;", " ", text, flags=re.I)

    text = re.sub(r"<[^>]+>|&lt;[^&]*&gt;", " ", text)
    text = FOOTNOTE_SYMBOL_RE.sub(" ", text)

    return re.sub(r"\s+", " ", text).strip()


def clean_match_text(value) -> str:
    text = normalize_text(value).lower()
    text = text.replace("&amp;", " and ")
    text = text.replace("&", " and ")

    text = re.sub(r"\bfund\s+of\s+funds?\b", " fof ", text)
    text = re.sub(r"\bfof\b", " fof ", text)

    text = re.sub(r"[^a-z0-9]+", " ", text)

    return re.sub(r"\s+", " ", text).strip()


def scheme_tokens(value) -> list[str]:
    return [
        token
        for token in clean_match_text(value).split()
        if token not in SCHEME_STOPWORDS and token != "and"
    ]


def canonical_scheme_key(value) -> str:
    return " ".join(sorted(set(scheme_tokens(value))))


def compact_text(value) -> str:
    return clean_match_text(value).replace(" ", "")


def scheme_match_score(excel_name: str, ppt_name: str) -> float:
    excel_clean = clean_match_text(excel_name)
    ppt_clean = clean_match_text(ppt_name)

    if not excel_clean or not ppt_clean:
        return 0.0

    if excel_clean == ppt_clean:
        return 1.0

    if compact_text(excel_name) == compact_text(ppt_name):
        return 1.0

    if excel_clean in ppt_clean or ppt_clean in excel_clean:
        return 0.98

    excel_tokens = set(scheme_tokens(excel_name))
    ppt_tokens = set(scheme_tokens(ppt_name))

    if not excel_tokens or not ppt_tokens:
        return 0.0

    if excel_tokens == ppt_tokens:
        return 1.0

    intersection = excel_tokens & ppt_tokens
    union = excel_tokens | ppt_tokens

    jaccard = len(intersection) / len(union)
    coverage = len(intersection) / min(len(excel_tokens), len(ppt_tokens))
    seq = SequenceMatcher(
        None,
        " ".join(sorted(excel_tokens)),
        " ".join(sorted(ppt_tokens)),
    ).ratio()

    if coverage >= 0.90 and jaccard >= 0.72:
        return max(coverage, seq)

    return max(jaccard, seq * 0.90)


def natural_key(text: str):
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", str(text))
    ]


def excel_engine(path: Path) -> str:
    return "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"


def parse_excel_date(value) -> Optional[date]:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        return value.date()

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = normalize_text(value)

    if not text:
        return None

    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)

    return None if pd.isna(parsed) else parsed.date()


def parse_user_date(value: str) -> date:
    text = normalize_text(value)

    if not text:
        return date.today()

    return pd.to_datetime(text, errors="raise", dayfirst=True).date()


def add_months(input_date: date, months: int) -> date:
    month_index = input_date.month - 1 + months
    year = input_date.year + month_index // 12
    month = month_index % 12 + 1

    month_days = [
        31,
        29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
        31,
        30,
        31,
        30,
        31,
        31,
        30,
        31,
        30,
        31,
    ]

    day = min(input_date.day, month_days[month - 1])

    return date(year, month, day)


def period_completion_date(
    launch_date: Optional[date],
    period: Period,
) -> Optional[date]:
    if launch_date is None:
        return None

    if period.unit == "day":
        return launch_date + timedelta(days=period.value)

    if period.unit == "month":
        return add_months(launch_date, period.value)

    try:
        return launch_date.replace(year=launch_date.year + period.value)
    except ValueError:
        return launch_date.replace(year=launch_date.year + period.value, day=28)


def completed_period(
    launch_date: Optional[date],
    as_of_date: date,
    period: Period,
) -> bool:
    completion = period_completion_date(launch_date, period)
    return completion is not None and as_of_date >= completion


def is_na_text(value: str) -> bool:
    text = normalize_text(value).strip().lower()
    compact = re.sub(r"[^a-z0-9]+", "", text)

    return compact in {
        "",
        "na",
        "n/a",
        "notapplicable",
        "nan",
        "nil",
    } or text == "-"


def looks_like_number_or_percent(value: str) -> bool:
    text = normalize_text(value)

    if is_na_text(text):
        return False

    return bool(re.search(r"[-+]?\d", text))


def summarize_report_dataframe(df: pd.DataFrame) -> tuple[int, int, int, int]:
    rows = len(df)
    fix_rows = 0
    ok_rows = 0
    review_rows = 0

    if "Result" in df.columns:
        result_series = df["Result"].astype(str).str.upper()
        fix_rows = int((result_series == "FIX").sum())
        ok_rows = int((result_series == "OK").sum())
        review_rows = int((result_series == "REVIEW").sum())

    return rows, fix_rows, ok_rows, review_rows


# =========================================================
# PPT HELPERS
# =========================================================
def run_baseline_value(run) -> int:
    try:
        r_pr = run._r.rPr

        if r_pr is None:
            return 0

        baseline = r_pr.get("baseline")

        if baseline is None:
            for key, value in r_pr.attrib.items():
                if str(key).endswith("baseline"):
                    baseline = value
                    break

        return int(baseline) if baseline not in [None, ""] else 0

    except Exception:
        return 0


def text_frame_text_without_superscript(text_frame) -> str:
    parts = []

    for paragraph in text_frame.paragraphs:
        para = []

        for run in paragraph.runs:
            if run_baseline_value(run) != 0:
                continue

            para.append(run.text or "")

        text = "".join(para).strip()

        if text:
            parts.append(text)

    return "\n".join(parts)


def cell_text(cell) -> str:
    text = text_frame_text_without_superscript(cell.text_frame)
    return normalize_text(text or cell.text)


def iter_shapes(shapes):
    for shape in shapes:
        yield shape

        if hasattr(shape, "shapes"):
            yield from iter_shapes(shape.shapes)


# =========================================================
# PERIOD / TABLE DETECTION
# =========================================================
def detect_period_columns(
    row_cells: list[str],
    allowed_periods: Optional[set[Period]] = None,
) -> dict[Period, list[int]]:
    columns: dict[Period, list[int]] = {}

    for index, value in enumerate(row_cells):
        text = normalize_text(value).lower()
        text = re.sub(r"\s+", " ", text).strip()

        found: list[Period] = []

        for match in re.finditer(r"\b(\d{1,3})\s*(day|days|d)\b", text):
            found.append(Period("day", int(match.group(1))))

        for match in re.finditer(r"\b(\d{1,2})\s*(month|months|m)\b", text):
            found.append(Period("month", int(match.group(1))))

        for match in re.finditer(r"\b(\d{1,2})\s*(year|years|yr|yrs|y)\b", text):
            found.append(Period("year", int(match.group(1))))

        # Performance Scheme Direct PPT may show headers only as 1, 3, 5, 10.
        # Treat plain numeric headers as years only for common year periods.
        compact_number = re.sub(r"[^0-9]", "", text)

        if compact_number and text == compact_number:
            number = int(compact_number)

            if number in {1, 3, 5, 7, 10}:
                found.append(Period("year", number))

        for period in found:
            if allowed_periods is not None and period not in allowed_periods:
                continue

            columns.setdefault(period, []).append(index)

    return dict(sorted(columns.items(), key=lambda item: item[0].sort_key))


def row_has_period_headers(
    row_cells: list[str],
    allowed_periods: Optional[set[Period]] = None,
) -> bool:
    return len(detect_period_columns(row_cells, allowed_periods)) >= 2


def table_has_period_headers(
    shape,
    allowed_periods: Optional[set[Period]] = None,
) -> bool:
    if not getattr(shape, "has_table", False):
        return False

    for row in shape.table.rows:
        if row_has_period_headers(
            [cell_text(cell) for cell in row.cells],
            allowed_periods,
        ):
            return True

    return False


def is_sip_header_row(row_cells: list[str]) -> bool:
    allowed = {Period(unit, value) for unit, value in SIP_PERIODS}
    columns = detect_period_columns(row_cells, allowed)
    first = normalize_text(row_cells[0]) if row_cells else ""

    return len(columns) >= 3 and bool(clean_match_text(first))


def looks_like_fund_title_row(row_cells: list[str]) -> bool:
    first = normalize_text(row_cells[0]) if row_cells else ""

    if not first:
        return False

    if row_has_period_headers(row_cells):
        return False

    return bool(clean_match_text(first)) and (
        "fund" in first.lower()
        or "bandhan" in first.lower()
        or len(scheme_tokens(first)) >= 2
    )


def find_scheme_data_row(
    table_rows: list[list[str]],
    header_row_idx: int,
    end_row_idx: int,
) -> Optional[int]:
    for idx in range(header_row_idx + 1, end_row_idx):
        first = clean_match_text(table_rows[idx][0] if table_rows[idx] else "")

        if first == "scheme":
            return idx

    return None


# =========================================================
# FILE AUTO DETECTION
# =========================================================
def is_excel_file(path: Path) -> bool:
    return path.suffix.lower() in {
        ".xlsx",
        ".xlsm",
        ".xls",
    } and not path.name.startswith("~$")


def score_excel_file(path: Path, launch_col_index: int) -> int:
    score = 0
    name = path.name.lower()

    if "fundmanager" in name or "fund_manager" in name or "managerdetails" in name:
        score += 70

    if "clientinput" in name or "exclusion" in name or "report" in name or "output" in name:
        score -= 50

    try:
        df = pd.read_excel(
            path,
            engine=excel_engine(path),
            header=None,
            nrows=30,
        )

        if df.shape[1] > launch_col_index:
            score += 90

        sample = " ".join(
            normalize_text(v).lower()
            for v in df.iloc[:10, :7].values.flatten()
        )

        if "scheme" in sample or "fund" in sample:
            score += 30

        if "launch" in sample:
            score += 30

    except Exception:
        score -= 100

    return score


def score_pptx_file(path: Path, mode: str) -> int:
    score = 0
    name = path.name.lower()

    allowed_sip = {
        Period(unit, value)
        for unit, value in SIP_PERIODS
    }

    allowed = allowed_sip if mode == "sip" else None

    if mode == "sip":
        if "sip" in name:
            score += 120

        if "performance" in name:
            score += 30

        if "fund" in name and "manager" in name and "sip" not in name:
            score -= 150

    else:
        if "performance of the schemes" in name:
            score += 200

        if "performance" in name:
            score += 70

        if "direct" in name:
            score += 80

        if "sip" in name:
            score -= 180

        if "fund" in name and "manager" in name:
            score -= 180

    try:
        prs = Presentation(str(path))
        has_table = False

        for slide in prs.slides:
            for shape in iter_shapes(slide.shapes):
                if table_has_period_headers(shape, allowed):
                    has_table = True
                    score += 300
                    break

            if has_table:
                break

        if not has_table:
            return -999

    except Exception:
        return -999

    return score


def auto_detect_files(folder: Path, mode: str) -> dict:
    launch_col = (
        EXCEL_LAUNCH_REGULAR_COL_INDEX
        if mode == "sip"
        else EXCEL_LAUNCH_DIRECT_COL_INDEX
    )

    files = [
        p
        for p in folder.iterdir()
        if p.is_file() and not p.name.startswith("~$")
    ]

    excel_files = [
        p
        for p in files
        if is_excel_file(p)
    ]

    pptx_files = [
        p
        for p in files
        if p.suffix.lower() == ".pptx"
    ]

    excel_candidates = sorted(
        [
            (p, score_excel_file(p, launch_col))
            for p in excel_files
        ],
        key=lambda item: (-item[1], natural_key(item[0].name)),
    )

    pptx_candidates = sorted(
        [
            (p, score_pptx_file(p, mode))
            for p in pptx_files
        ],
        key=lambda item: (-item[1], natural_key(item[0].name)),
    )

    return {
        "excel_files": excel_files,
        "pptx_files": pptx_files,
        "excel_candidates": excel_candidates,
        "pptx_candidates": pptx_candidates,
        "excel": excel_candidates[0][0]
        if excel_candidates and excel_candidates[0][1] > 0
        else None,
        "pptx": pptx_candidates[0][0]
        if pptx_candidates and pptx_candidates[0][1] > 0
        else None,
    }


# =========================================================
# EXCEL INPUT
# =========================================================
def read_scheme_launches_from_excel(
    excel_path: Path,
    launch_col_index: int,
    sheet_name: str | None = None,
) -> list[SchemeLaunch]:
    engine = excel_engine(excel_path)
    xls = pd.ExcelFile(excel_path, engine=engine)
    sheet = sheet_name or xls.sheet_names[0]

    df = pd.read_excel(
        excel_path,
        engine=engine,
        sheet_name=sheet,
        header=None,
    )

    if df.empty:
        raise ValueError("Excel sheet is empty.")

    if df.shape[1] <= launch_col_index:
        col_letter = "E" if launch_col_index == EXCEL_LAUNCH_REGULAR_COL_INDEX else "F"
        raise ValueError(f"Excel must have column {col_letter} for launch date.")

    rows: list[SchemeLaunch] = []

    for _, row in df.iterrows():
        scheme_name = normalize_text(row.iloc[EXCEL_SCHEME_COL_INDEX])

        if not scheme_name:
            continue

        if clean_match_text(scheme_name) in {
            "scheme",
            "scheme name",
            "fund",
            "fund name",
        }:
            continue

        launch_date = parse_excel_date(row.iloc[launch_col_index])

        rows.append(
            SchemeLaunch(
                scheme_name=scheme_name,
                launch_date=launch_date,
            )
        )

    return rows


# =========================================================
# PPT BLOCK EXTRACTION
# =========================================================
def extract_sip_blocks_from_ppt(
    pptx_path: Path,
    object_name_filter: str = "",
) -> list[TableBlock]:
    allowed = {
        Period(unit, value)
        for unit, value in SIP_PERIODS
    }

    prs = Presentation(str(pptx_path))
    blocks: list[TableBlock] = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        for shape in iter_shapes(slide.shapes):
            if not getattr(shape, "has_table", False):
                continue

            shape_name = getattr(shape, "name", "")

            if object_name_filter and shape_name != object_name_filter:
                continue

            table_rows = [
                [
                    cell_text(cell)
                    for cell in row.cells
                ]
                for row in shape.table.rows
            ]

            header_indices = [
                index
                for index, row_cells in enumerate(table_rows)
                if is_sip_header_row(row_cells)
            ]

            for header_pos, header_index in enumerate(header_indices):
                next_header = (
                    header_indices[header_pos + 1]
                    if header_pos + 1 < len(header_indices)
                    else len(table_rows)
                )

                header_cells = table_rows[header_index]
                period_columns = detect_period_columns(header_cells, allowed)
                scheme_text = header_cells[0]

                rows_by_period: dict[Period, list[str]] = {
                    period: []
                    for period in period_columns
                }

                for row_index in range(header_index + 1, next_header):
                    row_cells = table_rows[row_index]

                    for period, col_indexes in period_columns.items():
                        for col_index in col_indexes:
                            if col_index < len(row_cells):
                                value = normalize_text(row_cells[col_index])

                                if value:
                                    rows_by_period.setdefault(period, []).append(value)

                blocks.append(
                    TableBlock(
                        scheme_text=scheme_text,
                        slide=slide_index,
                        shape_name=shape_name,
                        start_row=header_index + 1,
                        end_row=next_header,
                        period_columns=period_columns,
                        rows_by_period=rows_by_period,
                    )
                )

    if not blocks:
        raise ValueError(
            "No SIP performance table blocks found. "
            "Expected scheme header row with 1 Year, 3 Years, 5 Years, etc."
        )

    return blocks


def extract_scheme_blocks_from_ppt(
    pptx_path: Path,
    object_name_filter: str = "",
) -> list[TableBlock]:
    prs = Presentation(str(pptx_path))
    blocks: list[TableBlock] = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        for shape in iter_shapes(slide.shapes):
            if not getattr(shape, "has_table", False):
                continue

            shape_name = getattr(shape, "name", "")

            if object_name_filter and shape_name != object_name_filter:
                continue

            table_rows = [
                [
                    cell_text(cell)
                    for cell in row.cells
                ]
                for row in shape.table.rows
            ]

            title_rows = [
                index
                for index, row_cells in enumerate(table_rows)
                if looks_like_fund_title_row(row_cells)
            ]

            for pos, title_idx in enumerate(title_rows):
                next_title = (
                    title_rows[pos + 1]
                    if pos + 1 < len(title_rows)
                    else len(table_rows)
                )

                scheme_text = table_rows[title_idx][0]

                best_header_idx = None
                best_cols: dict[Period, list[int]] = {}

                for scan_idx in range(title_idx + 1, min(next_title, title_idx + 12)):
                    cols = detect_period_columns(
                        table_rows[scan_idx],
                        allowed_periods=None,
                    )

                    if sum(len(v) for v in cols.values()) > sum(
                        len(v)
                        for v in best_cols.values()
                    ):
                        best_cols = cols
                        best_header_idx = scan_idx

                if best_header_idx is None or not best_cols:
                    continue

                scheme_row_idx = find_scheme_data_row(
                    table_rows,
                    best_header_idx,
                    next_title,
                )

                if scheme_row_idx is None:
                    continue

                rows_by_period: dict[Period, list[str]] = {
                    period: []
                    for period in best_cols
                }

                scheme_row = table_rows[scheme_row_idx]

                for period, col_indexes in best_cols.items():
                    for col_idx in col_indexes:
                        if col_idx < len(scheme_row):
                            value = normalize_text(scheme_row[col_idx])

                            if value:
                                rows_by_period.setdefault(period, []).append(value)

                blocks.append(
                    TableBlock(
                        scheme_text=scheme_text,
                        slide=slide_index,
                        shape_name=shape_name,
                        start_row=title_idx + 1,
                        end_row=next_title,
                        period_columns=best_cols,
                        rows_by_period=rows_by_period,
                    )
                )

    if not blocks:
        raise ValueError(
            "No Performance Scheme blocks found. "
            "Expected fund title row, period header row, and Scheme data row."
        )

    return blocks


def find_block_for_scheme(
    blocks: list[TableBlock],
    scheme_name: str,
    threshold: float = 0.82,
) -> tuple[Optional[TableBlock], float, str]:
    best_block = None
    best_score = 0.0
    best_reason = ""
    excel_key = canonical_scheme_key(scheme_name)

    for block in blocks:
        score = scheme_match_score(scheme_name, block.scheme_text)

        if score > best_score:
            best_score = score
            best_block = block
            best_reason = (
                f"Excel key='{excel_key}' | "
                f"PPT key='{canonical_scheme_key(block.scheme_text)}'"
            )

    if best_score >= threshold:
        return best_block, best_score, best_reason

    return None, best_score, best_reason


def period_column_has_data(values: list[str]) -> bool:
    return any(
        looks_like_number_or_percent(value)
        for value in values
    )


def period_column_all_na_or_blank(values: list[str]) -> bool:
    if not values:
        return True

    return all(
        is_na_text(value)
        for value in values
    )


# =========================================================
# REPORT BUILDING / WRITING
# =========================================================
def build_report_rows(
    excel_path: Path,
    pptx_path: Path,
    as_of_date: date,
    mode: str,
    sheet_name: str | None = None,
    object_name_filter: str = "",
) -> pd.DataFrame:
    launch_col = (
        EXCEL_LAUNCH_REGULAR_COL_INDEX
        if mode == "sip"
        else EXCEL_LAUNCH_DIRECT_COL_INDEX
    )

    plan_name = "Regular" if mode == "sip" else "Direct"
    launch_col_name = "Launch Date Regular" if mode == "sip" else "Launch Date Direct"

    launches = read_scheme_launches_from_excel(
        excel_path=excel_path,
        launch_col_index=launch_col,
        sheet_name=sheet_name,
    )

    blocks = (
        extract_sip_blocks_from_ppt(
            pptx_path,
            object_name_filter,
        )
        if mode == "sip"
        else extract_scheme_blocks_from_ppt(
            pptx_path,
            object_name_filter,
        )
    )

    rows = []

    for item in launches:
        block, match_score, match_reason = find_block_for_scheme(
            blocks,
            item.scheme_name,
        )

        if not block:
            rows.append(
                {
                    "Result": "FIX",
                    "Plan Checked": plan_name,
                    "Scheme Name": item.scheme_name,
                    launch_col_name: (
                        item.launch_date.strftime("%d-%b-%Y")
                        if item.launch_date
                        else "Missing/Invalid"
                    ),
                    "As Of Date": as_of_date.strftime("%d-%b-%Y"),
                    "Period": "Scheme Not Found",
                    "Period Completion Date": "",
                    "Expected": "Review",
                    "Actual": "Scheme Not Found",
                    "PPT Location": "Not found",
                    "Matched PPT Scheme Text": "",
                    "Scheme Match Score": round(match_score, 3),
                    "Scheme Match Reason": match_reason,
                    "Column Values Read": "",
                    "Action Required": (
                        "No matching performance block was found in PPT for this scheme."
                    ),
                }
            )
            continue

        periods = sorted(
            block.period_columns.keys(),
            key=lambda p: p.sort_key,
        )

        for period in periods:
            eligible = completed_period(
                item.launch_date,
                as_of_date,
                period,
            )

            completion = period_completion_date(
                item.launch_date,
                period,
            )

            values = block.rows_by_period.get(period, [])

            has_data = period_column_has_data(values)
            all_na = period_column_all_na_or_blank(values)

            location = (
                f"Slide {block.slide}, Shape '{block.shape_name}', "
                f"Rows {block.start_row}-{block.end_row}, "
                f"Column {period.label}"
            )

            if eligible and has_data:
                result = "OK"
                actual = "Figures Found"
                action = (
                    "No action required. Scheme has completed the period "
                    "and figures are present."
                )

            elif eligible and not has_data:
                result = "FIX"
                actual = "NA/Blank"
                action = (
                    f"Scheme has completed {period.label}, "
                    f"so figures should be present instead of NA/blank."
                )

            elif not eligible and all_na:
                result = "OK"
                actual = "NA/Blank"
                action = (
                    "No action required. Scheme has not completed the period "
                    "and column is NA/blank."
                )

            else:
                result = "FIX"
                actual = "Figures Found"
                action = (
                    f"Scheme has not completed {period.label}, "
                    f"so this column should be NA."
                )

            rows.append(
                {
                    "Result": result,
                    "Plan Checked": plan_name,
                    "Scheme Name": item.scheme_name,
                    launch_col_name: (
                        item.launch_date.strftime("%d-%b-%Y")
                        if item.launch_date
                        else "Missing/Invalid"
                    ),
                    "As Of Date": as_of_date.strftime("%d-%b-%Y"),
                    "Period": period.label,
                    "Period Completion Date": (
                        completion.strftime("%d-%b-%Y")
                        if completion
                        else ""
                    ),
                    "Expected": "Figures" if eligible else "NA",
                    "Actual": actual,
                    "PPT Location": location,
                    "Matched PPT Scheme Text": block.scheme_text,
                    "Scheme Match Score": round(match_score, 3),
                    "Scheme Match Reason": match_reason,
                    "Column Values Read": "; ".join(values),
                    "Action Required": action,
                }
            )

    return pd.DataFrame(rows)


def write_report(
    df: pd.DataFrame,
    output_path: Path,
    sheet_name: str,
):
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")

    temp_path = output_path.with_name(output_path.stem + "_temp.xlsx")

    if temp_path.exists():
        temp_path.unlink()

    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    wb = load_workbook(temp_path)
    ws = wb[sheet_name]

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    dark_blue = "1F4E78"
    white = "FFFFFF"

    green_completed = "C6EFCE"
    blue_not_completed = "DDEBF7"
    red_fix = "FFC7CE"
    yellow_review = "FFF2CC"

    thin = Side(style="thin", color="D9D9D9")
    medium = Side(style="medium", color="808080")

    for cell in ws[1]:
        cell.font = Font(bold=True, color=white, size=11)
        cell.fill = PatternFill("solid", fgColor=dark_blue)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        cell.border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin,
        )

    header_map = {
        str(cell.value or "").strip(): cell.column
        for cell in ws[1]
    }

    result_col = header_map.get("Result", 1)
    scheme_col = header_map.get("Scheme Name", 3)
    expected_col = header_map.get("Expected", 8)
    actual_col = header_map.get("Actual", 9)

    previous_scheme = None

    for row_idx in range(2, ws.max_row + 1):
        result = str(ws.cell(row_idx, result_col).value or "").upper().strip()
        scheme = str(ws.cell(row_idx, scheme_col).value or "").strip()
        expected = str(ws.cell(row_idx, expected_col).value or "").upper().strip()

        if result == "FIX":
            fill = red_fix
        elif expected == "FIGURES":
            fill = green_completed
        elif expected == "NA":
            fill = blue_not_completed
        else:
            fill = yellow_review

        is_new_scheme = previous_scheme is not None and scheme != previous_scheme
        previous_scheme = scheme

        for cell in ws[row_idx]:
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(
                left=thin,
                right=thin,
                top=medium if is_new_scheme else thin,
                bottom=thin,
            )

        ws.cell(row_idx, result_col).font = Font(bold=True)
        ws.cell(row_idx, expected_col).font = Font(bold=True)
        ws.cell(row_idx, actual_col).font = Font(bold=True)

    widths = {
        "A": 10,
        "B": 14,
        "C": 58,
        "D": 20,
        "E": 18,
        "F": 18,
        "G": 22,
        "H": 14,
        "I": 24,
        "J": 48,
        "K": 62,
        "L": 18,
        "M": 70,
        "N": 85,
        "O": 95,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    legend_col = ws.max_column + 2

    ws.cell(1, legend_col).value = "Colour Legend"
    ws.cell(1, legend_col).font = Font(bold=True, color=white)
    ws.cell(1, legend_col).fill = PatternFill("solid", fgColor=dark_blue)

    legend_items = [
        ("Completed period: figures expected/found", green_completed),
        ("Not completed period: NA expected/found", blue_not_completed),
        ("FIX required", red_fix),
    ]

    for idx, (label, colour) in enumerate(legend_items, start=2):
        cell = ws.cell(idx, legend_col)
        cell.value = label
        cell.fill = PatternFill("solid", fgColor=colour)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin,
        )

    ws.column_dimensions[ws.cell(1, legend_col).column_letter].width = 46
    ws.auto_filter.ref = ws.dimensions

    wb.save(temp_path)

    try:
        if output_path.exists():
            output_path.unlink()

        temp_path.replace(output_path)

    except PermissionError:
        raise PermissionError(
            f"Cannot replace report because it is open in Excel:\n"
            f"{output_path}\n\n"
            f"Close it and run again."
        )


def run_backend(
    excel_path,
    pptx_path,
    output_path,
    as_of_date_text,
    sheet_name,
    object_name_filter,
    mode: str,
    progress_cb=None,
) -> CheckResult:
    def log(message: str):
        if progress_cb:
            progress_cb(message)

    label = "SIP" if mode == "sip" else "PerformanceScheme"

    excel_path = Path(excel_path)
    pptx_path = Path(pptx_path)
    output_path = Path(output_path)

    log(f"[{label}] Progress: 0%")
    log(f"[{label}] Excel: {excel_path}")
    log(f"[{label}] PPTX: {pptx_path}")
    log(f"[{label}] Output: {output_path}")

    as_of_date = parse_user_date(as_of_date_text)

    log(f"[{label}] As-of Date: {as_of_date.strftime('%d-%m-%Y')}")
    log(f"[{label}] Building report rows...")

    df = build_report_rows(
        excel_path=excel_path,
        pptx_path=pptx_path,
        as_of_date=as_of_date,
        mode=mode,
        sheet_name=sheet_name.strip() or None,
        object_name_filter=object_name_filter.strip(),
    )

    log(f"[{label}] Progress: 75%")

    write_report(
        df=df,
        output_path=output_path,
        sheet_name="SIP Tenure Check" if mode == "sip" else "Performance Check",
    )

    rows, fix_rows, ok_rows, review_rows = summarize_report_dataframe(df)

    log(f"[{label}] Rows: {rows}")
    log(f"[{label}] FIX rows: {fix_rows}")
    log(f"[{label}] OK rows: {ok_rows}")
    log(f"[{label}] REVIEW rows: {review_rows}")
    log(f"[{label}] Saved: {output_path}")
    log(f"[{label}] Progress: 100%")

    return CheckResult(
        output_path=output_path,
        rows=rows,
        fix_rows=fix_rows,
        ok_rows=ok_rows,
        review_rows=review_rows,
    )


# =========================================================
# GUI
# =========================================================
class CombinedChecksWindow(QWidget):
    def __init__(self, state: Any | None = None):
        super().__init__()

        self.state = state
        self.thread_pool = QThreadPool.globalInstance()

        self.sip_worker = None
        self.scheme_worker = None
        self.default_output_folder = self._resolve_default_output_folder()

        self._build_ui()
        self._apply_default_output_paths()
        self._set_sip_idle()
        self._set_scheme_idle()


    def _app_run_folder(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parents[1]

    def _resolve_default_output_folder(self) -> Path:
        state_folder = getattr(self.state, "tab4_output_folder", None) if self.state is not None else None
        if state_folder:
            folder = Path(state_folder)
        else:
            output_root = getattr(self.state, "output_root", None) if self.state is not None else None
            folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "04_sip_performance_check"
        folder.mkdir(parents=True, exist_ok=True)
        if self.state is not None:
            setattr(self.state, "tab4_output_folder", folder)
        return folder

    def _default_output_path(self, default_name: str) -> Path:
        return self._resolve_default_output_folder() / default_name

    def _apply_default_output_paths(self) -> None:
        if hasattr(self, "sip_output_edit") and not self.sip_output_edit.text().strip():
            self.sip_output_edit.setText(str(self._default_output_path("sip_performance_tenure_check_report.xlsx")))
        if hasattr(self, "scheme_output_edit") and not self.scheme_output_edit.text().strip():
            self.scheme_output_edit.setText(str(self._default_output_path("performance_scheme_direct_check_report.xlsx")))

    def _build_ui(self):
        self.setWindowTitle(APP_TITLE)

        root = QHBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(8)

        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)

        root.addWidget(splitter)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(6, 6, 6, 6)

        log_title = QLabel("Process Logs")
        log_title.setStyleSheet("font-weight: bold;")

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        self.clear_log_btn = QPushButton("Clear Log")

        left_layout.addWidget(log_title)
        left_layout.addWidget(self.log_box, stretch=1)
        left_layout.addWidget(self.clear_log_btn)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        panel = QWidget()
        right_layout = QVBoxLayout(panel)
        right_layout.setContentsMargins(6, 6, 12, 6)
        right_layout.setSpacing(10)

        title = QLabel("Combined SIP Performance + Performance Scheme Checks")
        title.setStyleSheet("font-weight: bold; font-size: 16px;")

        desc = QLabel(
            "SIP checks Regular launch date column E against 1/3/5/7/10 year SIP columns. "
            "Performance Scheme Direct checks Direct launch date column F against periods actually present in PPT: "
            "years, days, and months. Scheme matching treats FoF/Fund of Funds and Regular/Direct/Plan/Growth "
            "differences as equivalent."
        )
        desc.setWordWrap(True)

        right_layout.addWidget(title)
        right_layout.addWidget(desc)
        right_layout.addWidget(self._build_section("sip"))
        right_layout.addWidget(self._build_section("scheme"))
        right_layout.addStretch(1)

        scroll.setWidget(panel)

        splitter.addWidget(left)
        splitter.addWidget(scroll)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setSizes([360, 1080])

        self.clear_log_btn.clicked.connect(self.log_box.clear)

    def _build_section(self, mode: str):
        is_sip = mode == "sip"
        prefix = "sip" if is_sip else "scheme"

        group = QGroupBox(
            "Section 1 - SIP Performance Check"
            if is_sip
            else "Section 2 - Performance Scheme Check Direct"
        )

        layout = QVBoxLayout(group)

        info = QLabel(
            "Regular launch date column E is checked against SIP performance tenures: "
            "1/3/5/7/10 Years."
            if is_sip
            else
            "Direct launch date column F is checked against Performance Scheme periods found in PPT, "
            "e.g. 7/15/30 Days, 6 Months, 1/3/5/10 Years."
        )
        info.setWordWrap(True)

        layout.addWidget(info)

        folder_group = QGroupBox("1. Folder Auto Detect")
        folder_grid = QGridLayout(folder_group)

        folder_edit = QLineEdit()
        browse_folder_btn = QPushButton("Browse Folder...")
        find_files_btn = QPushButton("Find Files")

        folder_grid.addWidget(QLabel("Source Folder:"), 0, 0)
        folder_grid.addWidget(folder_edit, 0, 1)
        folder_grid.addWidget(browse_folder_btn, 0, 2)
        folder_grid.addWidget(find_files_btn, 0, 3)

        layout.addWidget(folder_group)

        input_group = QGroupBox("2. Input Files")
        input_grid = QGridLayout(input_group)

        excel_edit = QLineEdit()
        pptx_edit = QLineEdit()

        browse_excel_btn = QPushButton("Browse Excel...")
        browse_pptx_btn = QPushButton("Browse PPTX...")

        input_grid.addWidget(QLabel("fundmanagerdetails Excel:"), 0, 0)
        input_grid.addWidget(excel_edit, 0, 1)
        input_grid.addWidget(browse_excel_btn, 0, 2)

        input_grid.addWidget(
            QLabel("SIP Performance PPTX:" if is_sip else "Performance Scheme PPTX:"),
            1,
            0,
        )
        input_grid.addWidget(pptx_edit, 1, 1)
        input_grid.addWidget(browse_pptx_btn, 1, 2)

        layout.addWidget(input_group)

        settings_group = QGroupBox("3. Settings")
        settings_grid = QGridLayout(settings_group)

        asof_edit = QLineEdit(date.today().strftime("%d-%m-%Y"))
        sheet_edit = QLineEdit()
        object_filter_edit = QLineEdit()

        sheet_edit.setPlaceholderText("Optional. Blank = first sheet")
        object_filter_edit.setPlaceholderText("Optional PPT table/object name filter")

        settings_grid.addWidget(QLabel("As-of Date:"), 0, 0)
        settings_grid.addWidget(asof_edit, 0, 1)

        settings_grid.addWidget(QLabel("Excel Sheet Name:"), 1, 0)
        settings_grid.addWidget(sheet_edit, 1, 1)

        settings_grid.addWidget(QLabel("PPT Object Name Filter:"), 2, 0)
        settings_grid.addWidget(object_filter_edit, 2, 1)

        layout.addWidget(settings_group)

        output_group = QGroupBox("4. Output")
        output_grid = QGridLayout(output_group)

        output_edit = QLineEdit()
        browse_output_btn = QPushButton("Save Report As...")

        default_name = (
            "sip_performance_tenure_check_report.xlsx"
            if is_sip
            else "performance_scheme_direct_check_report.xlsx"
        )

        output_grid.addWidget(QLabel("Output Report XLSX:"), 0, 0)
        output_grid.addWidget(output_edit, 0, 1)
        output_grid.addWidget(browse_output_btn, 0, 2)

        layout.addWidget(output_group)

        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)

        status_label = QLabel("Idle")

        progress_bar = QProgressBar()
        progress_bar.setRange(0, 100)
        progress_bar.setValue(0)

        progress_layout.addWidget(status_label)
        progress_layout.addWidget(progress_bar)

        layout.addWidget(progress_group)

        run_row = QHBoxLayout()
        run_row.addStretch(1)

        run_btn = QPushButton(
            "Run SIP Performance Check"
            if is_sip
            else "Run Performance Scheme Check Direct"
        )

        run_row.addWidget(run_btn)
        layout.addLayout(run_row)

        setattr(self, f"{prefix}_folder_edit", folder_edit)
        setattr(self, f"{prefix}_excel_edit", excel_edit)
        setattr(self, f"{prefix}_pptx_edit", pptx_edit)
        setattr(self, f"{prefix}_asof_edit", asof_edit)
        setattr(self, f"{prefix}_sheet_edit", sheet_edit)
        setattr(self, f"{prefix}_object_filter_edit", object_filter_edit)
        setattr(self, f"{prefix}_output_edit", output_edit)
        setattr(self, f"{prefix}_status_label", status_label)
        setattr(self, f"{prefix}_progress_bar", progress_bar)
        setattr(self, f"{prefix}_run_btn", run_btn)

        browse_folder_btn.clicked.connect(lambda: self.browse_folder(mode))
        find_files_btn.clicked.connect(lambda: self.find_files(mode))
        browse_excel_btn.clicked.connect(lambda: self.browse_excel(excel_edit))
        browse_pptx_btn.clicked.connect(
            lambda: self.browse_pptx(
                pptx_edit,
                output_edit,
                default_name,
            )
        )
        browse_output_btn.clicked.connect(
            lambda: self.browse_output_xlsx(
                output_edit,
                default_name,
            )
        )
        run_btn.clicked.connect(lambda: self.start_process(mode))

        return group

    def log(self, message: str):
        self.log_box.append(str(message))
        self.log_box.ensureCursorVisible()

    def _set_sip_idle(self):
        self.sip_progress_bar.setRange(0, 100)
        self.sip_progress_bar.setValue(0)
        self.sip_status_label.setText("Idle")

    def _set_scheme_idle(self):
        self.scheme_progress_bar.setRange(0, 100)
        self.scheme_progress_bar.setValue(0)
        self.scheme_status_label.setText("Idle")

    def on_progress(self, payload, bar, label):
        text = str(payload)
        match = re.search(r"(\d{1,3})\s*%", text)

        if match:
            value = max(0, min(100, int(match.group(1))))
            bar.setRange(0, 100)
            bar.setValue(value)
            label.setText(f"{value}% complete")

        self.log(text)

    def browse_excel(self, target):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel",
            str(Path.cwd()),
            "Excel Files (*.xlsx *.xlsm *.xls)",
        )

        if path:
            target.setText(path)

    def browse_pptx(self, target, output_target, default_name):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select PPTX",
            str(Path.cwd()),
            "PowerPoint Files (*.pptx)",
        )

        if path:
            target.setText(path)

            if not output_target.text().strip():
                output_target.setText(
                    str(self._default_output_path(default_name))
                )

    def browse_output_xlsx(self, target, default_name):
        default_path = target.text().strip() or str(self._default_output_path(default_name))

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report As",
            default_path,
            "Excel Files (*.xlsx)",
        )

        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"

            target.setText(path)

    def browse_folder(self, mode: str):
        folder_edit = self.sip_folder_edit if mode == "sip" else self.scheme_folder_edit

        title = (
            "Select SIP Source Folder"
            if mode == "sip"
            else "Select Performance Scheme Source Folder"
        )

        folder = QFileDialog.getExistingDirectory(
            self,
            title,
            folder_edit.text().strip() or str(Path.cwd()),
        )

        if folder:
            folder_edit.setText(folder)
            self.find_files(mode)

    def find_files(self, mode: str):
        prefix = "sip" if mode == "sip" else "scheme"
        folder = Path(getattr(self, f"{prefix}_folder_edit").text().strip())

        if not folder.is_dir():
            QMessageBox.warning(
                self,
                "Invalid Folder",
                "Please select a valid source folder.",
            )
            return

        detected = auto_detect_files(folder, mode)

        excel_edit = getattr(self, f"{prefix}_excel_edit")
        pptx_edit = getattr(self, f"{prefix}_pptx_edit")
        output_edit = getattr(self, f"{prefix}_output_edit")

        default_name = (
            "sip_performance_tenure_check_report.xlsx"
            if mode == "sip"
            else "performance_scheme_direct_check_report.xlsx"
        )

        label = "SIP" if mode == "sip" else "PerformanceScheme"

        if detected.get("excel"):
            excel_edit.setText(str(detected["excel"]))

        if detected.get("pptx"):
            pptx_edit.setText(str(detected["pptx"]))

        if not output_edit.text().strip():
            output_edit.setText(str(self._default_output_path(default_name)))

        self.log(f"[{label}] Auto-detect Excel: {detected.get('excel') or 'Not found'}")
        self.log(f"[{label}] Auto-detect PPTX: {detected.get('pptx') or 'Not found'}")

        for path, score in detected.get("pptx_candidates", []):
            self.log(f"[{label}] PPTX candidate: {path.name} | score: {score}")

    def validate_common_inputs(self, mode: str):
        prefix = "sip" if mode == "sip" else "scheme"
        label = "SIP" if mode == "sip" else "Performance Scheme"

        excel = getattr(self, f"{prefix}_excel_edit").text().strip()
        pptx = getattr(self, f"{prefix}_pptx_edit").text().strip()
        output = getattr(self, f"{prefix}_output_edit").text().strip()
        asof = getattr(self, f"{prefix}_asof_edit").text().strip()

        if not excel or not Path(excel).exists():
            QMessageBox.warning(
                self,
                "Invalid Input",
                f"{label}: please select a valid Excel file.",
            )
            return None

        if not pptx or not Path(pptx).exists():
            QMessageBox.warning(
                self,
                "Invalid Input",
                f"{label}: please select a valid PPTX file.",
            )
            return None

        if not output:
            QMessageBox.warning(
                self,
                "Invalid Output",
                f"{label}: please select output report path.",
            )
            return None

        if not output.lower().endswith(".xlsx"):
            output += ".xlsx"

        if not asof:
            QMessageBox.warning(
                self,
                "Invalid Date",
                f"{label}: please enter As-of Date.",
            )
            return None

        return Path(excel), Path(pptx), Path(output), asof

    def start_process(self, mode: str):
        validated = self.validate_common_inputs(mode)

        if validated is None:
            return

        excel, pptx, output, asof = validated

        prefix = "sip" if mode == "sip" else "scheme"
        label = "SIP" if mode == "sip" else "PerformanceScheme"

        self.log("=" * 80)
        self.log(f"[{label}] Starting check...")

        run_btn = getattr(self, f"{prefix}_run_btn")
        bar = getattr(self, f"{prefix}_progress_bar")
        status = getattr(self, f"{prefix}_status_label")

        run_btn.setEnabled(False)
        bar.setRange(0, 0)
        status.setText("Running...")

        sheet = getattr(self, f"{prefix}_sheet_edit").text()
        obj = getattr(self, f"{prefix}_object_filter_edit").text()

        worker = FunctionWorker(
            run_backend,
            excel,
            pptx,
            output,
            asof,
            sheet,
            obj,
            mode,
        )

        if mode == "sip":
            self.sip_worker = worker
        else:
            self.scheme_worker = worker

        worker.signals.progress.connect(
            lambda payload: self.on_progress(
                payload,
                bar,
                status,
            )
        )
        worker.signals.finished.connect(
            lambda result: self.on_finished(mode, result)
        )
        worker.signals.error.connect(
            lambda error_text: self.on_error(mode, error_text)
        )

        self.thread_pool.start(worker)

    def on_finished(self, mode: str, result: CheckResult):
        prefix = "sip" if mode == "sip" else "scheme"

        title = (
            "SIP Performance Check Complete"
            if mode == "sip"
            else "Performance Scheme Check Complete"
        )

        if mode == "sip":
            self.sip_worker = None
        else:
            self.scheme_worker = None

        getattr(self, f"{prefix}_run_btn").setEnabled(True)

        bar = getattr(self, f"{prefix}_progress_bar")
        status = getattr(self, f"{prefix}_status_label")

        bar.setRange(0, 100)
        bar.setValue(100)
        status.setText("Completed")

        QMessageBox.information(
            self,
            title,
            f"Report saved at:\n{result.output_path}\n\n"
            f"Rows: {result.rows}\n"
            f"FIX: {result.fix_rows}\n"
            f"OK: {result.ok_rows}\n"
            f"REVIEW: {result.review_rows}",
        )

    def on_error(self, mode: str, error_text: str):
        prefix = "sip" if mode == "sip" else "scheme"

        title = (
            "SIP Performance Check Failed"
            if mode == "sip"
            else "Performance Scheme Check Failed"
        )

        label = "SIP" if mode == "sip" else "PerformanceScheme"

        if mode == "sip":
            self.sip_worker = None
            self._set_sip_idle()
        else:
            self.scheme_worker = None
            self._set_scheme_idle()

        getattr(self, f"{prefix}_run_btn").setEnabled(True)
        getattr(self, f"{prefix}_status_label").setText("Failed")

        self.log(f"[{label}] ERROR")
        self.log(error_text)

        QMessageBox.critical(
            self,
            title,
            error_text,
        )


# Friendly aliases for dynamic tab loaders.

# =========================================================
# Tab 5-style Section Tabs GUI override
# ---------------------------------------------------------
# IMPORTANT:
# - Original code above is kept intact.
# - Processing/backend logic is not changed.
# - This block only overrides the GUI layout at runtime so the screen is split
#   into Section 1 and Section 2 tabs, like Tab 5.
# =========================================================

def _tab4_build_tab5_style_section_tab(self, log_box: QTextEdit, clear_handler: Callable[[], None], section_widget: QWidget) -> QWidget:
    tab = QWidget()
    main_layout = QHBoxLayout(tab)
    main_layout.setContentsMargins(6, 6, 6, 6)
    main_layout.setSpacing(8)

    left_widget = QWidget()
    left_layout = QVBoxLayout(left_widget)
    left_layout.setContentsMargins(6, 6, 6, 6)
    left_layout.setSpacing(8)

    log_title = QLabel("Process Logs")
    log_title.setStyleSheet("font-weight: bold; font-size: 13px;")
    left_layout.addWidget(log_title)

    log_box.setReadOnly(True)
    left_layout.addWidget(log_box, stretch=1)

    clear_btn = QPushButton("Clear Log")
    clear_btn.clicked.connect(clear_handler)
    left_layout.addWidget(clear_btn)

    left_widget.setMinimumWidth(380)
    left_widget.setMaximumWidth(420)

    scroll = QScrollArea()
    scroll.setWidgetResizable(True)

    panel = QWidget()
    right_layout = QVBoxLayout(panel)
    right_layout.setContentsMargins(6, 6, 12, 6)
    right_layout.setSpacing(10)
    right_layout.addWidget(section_widget)
    right_layout.addStretch(1)
    scroll.setWidget(panel)

    main_layout.addWidget(left_widget)
    main_layout.addWidget(scroll, stretch=1)
    return tab


def _tab4_tab5_style_build_ui(self):
    """Build GUI like Tab 5: separate Section 1 and Section 2 tabs."""
    self.setWindowTitle(APP_TITLE)

    root = QVBoxLayout(self)
    root.setContentsMargins(6, 6, 6, 6)
    root.setSpacing(8)

    title = QLabel("Combined SIP Performance + Performance Scheme Checks")
    title.setStyleSheet("font-weight: bold; font-size: 18px;")
    root.addWidget(title)

    desc = QLabel(
        "Section 1 checks SIP Performance tenures using Regular launch date column E. "
        "Section 2 checks Performance Scheme periods using Direct launch date column F."
    )
    desc.setWordWrap(True)
    root.addWidget(desc)

    self.section_tabs = QTabWidget()
    root.addWidget(self.section_tabs, stretch=1)

    self.sip_log_box = QTextEdit()
    self.sip_log_box.setReadOnly(True)
    sip_section = self._build_section("sip")
    sip_tab = self._build_tab5_style_section_tab(
        log_box=self.sip_log_box,
        clear_handler=self.sip_log_box.clear,
        section_widget=sip_section,
    )
    self.section_tabs.addTab(sip_tab, "Section 1 - SIP Performance Check")

    self.scheme_log_box = QTextEdit()
    self.scheme_log_box.setReadOnly(True)
    scheme_section = self._build_section("scheme")
    scheme_tab = self._build_tab5_style_section_tab(
        log_box=self.scheme_log_box,
        clear_handler=self.scheme_log_box.clear,
        section_widget=scheme_section,
    )
    self.section_tabs.addTab(scheme_tab, "Section 2 - Performance Scheme Check")

    # Backward-compatible alias for any old handler expecting self.log_box.
    self.log_box = self.sip_log_box


def _tab4_tab5_style_log(self, text: str) -> None:
    """Send log messages to the active section log box while keeping compatibility."""
    message = str(text)

    targets = []
    if hasattr(self, "section_tabs"):
        current_index = self.section_tabs.currentIndex()
        if current_index == 0 and hasattr(self, "sip_log_box"):
            targets.append(self.sip_log_box)
        elif current_index == 1 and hasattr(self, "scheme_log_box"):
            targets.append(self.scheme_log_box)

    if not targets:
        if hasattr(self, "sip_log_box"):
            targets.append(self.sip_log_box)
        if hasattr(self, "scheme_log_box"):
            targets.append(self.scheme_log_box)
        if not targets and hasattr(self, "log_box"):
            targets.append(self.log_box)

    for box in targets:
        box.append(message)
        box.ensureCursorVisible()


# Runtime override only. Original methods are preserved above.
CombinedChecksWindow._build_tab5_style_section_tab = _tab4_build_tab5_style_section_tab
CombinedChecksWindow._build_ui = _tab4_tab5_style_build_ui
CombinedChecksWindow.log = _tab4_tab5_style_log

CombinedChecksWidget = CombinedChecksWindow
Tab4Widget = CombinedChecksWindow


def main():
    app = QApplication(sys.argv)

    window = CombinedChecksWindow()
    window.resize(1450, 900)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()