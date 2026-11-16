from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import os
import re
import sys
import tempfile
import traceback
import unicodedata
from dataclasses import dataclass
from typing import Any, Callable

from PyQt6.QtCore import QObject, QRunnable, QThreadPool, Qt, QRect, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QColor, QPen
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QColorDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStyle,
    QStyleOptionViewItem,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.util import Pt


APP_TITLE = "Update Index / Scheme References + Page Numbers"


# =========================================================
# CONFIG
# =========================================================
SHAPE_NAME = "PageNumber"
SHAPE_NAME_ALIASES = {"PageNumber", "Page Number"}

INDEX_TABLE_OBJECT_NAME = "IndexTable"

INDEX_SEARCH_COL = 0
INDEX_PAGE_NO_COL = 1

FUNDMANAGER_SEARCH_COL = 1
FUNDMANAGER_PAGE_NO_COL = 2

LANDSCAPE_POS = {
    "left": 10233992,
    "top": 7154491,
    "width": 93980,
    "height": 151323,
}

PORTRAIT_POS = {
    "left": 7225120,
    "top": 10408324,
    "width": 198000,
    "height": 165100,
}

FONT_NAME_DARK = "ubuntu"
FONT_NAME_LIGHT = "Ubuntu"

FONT_SIZE = Pt(8)
FONT_BOLD = False
FONT_ITALIC = False

FONT_COLOR = RGBColor(35, 31, 32)
FONT_COLOR_LIGHT = RGBColor(254, 254, 254)

DEFAULT_FONT_NAME = "Ubuntu"
DEFAULT_FONT_SIZE = Pt(7)

PARAGRAPH_ALIGNMENT = PP_ALIGN.LEFT
VERTICAL_ALIGNMENT = MSO_VERTICAL_ANCHOR.MIDDLE

PPT_EXTENSIONS = (".pptx",)

OBJECTS_TO_SEARCH_IN = [
    "Fund Name",
    "object 18",
    "object 31",
    "object 9",
    "object 23",
    "object 4",
    "DateAsOn",
]

BLACKLIST_NAMES = set()

EXCLUDE_CURRENT_UPDATE_PPT = True

IGNORED_OBJECT_NAME_CHARS = ["ß", "¥"]

# Index page search aliases.
# If the text in IndexTable column 1 is on the left, the tool will also search
# the words listed on the right inside configured PPT objects.
INDEX_SEARCH_ALIASES = {
    "Snapshots": [
        "Equity Snapshot",
        "Snapshot",
        "Snapshots",
    ],
    "Snapshot": [
        "Equity Snapshot",
        "Snapshot",
        "Snapshots",
    ],
    "Commentary - Equity Outlook": [
        "Commentary - Equity Outlook",
        "Commentry - Equity Outlook",
        "Equity Outlook",
    ],
    "Commentry - Equity Outlook": [
        "Commentary - Equity Outlook",
        "Commentry - Equity Outlook",
        "Equity Outlook",
    ],
    "Commentary - Debt Outlook": [
        "Commentary - Debt Outlook",
        "Commentry - Debt Outlook",
        "Debt Outlook",
    ],
    "Commentry - Debt Outlook": [
        "Commentary - Debt Outlook",
        "Commentry - Debt Outlook",
        "Debt Outlook",
    ],
}


# =========================================================
# DATA MODELS
# =========================================================
@dataclass
class FileMeta:
    file_name: str
    file_stem: str
    path: str
    slides: int
    start_page: int
    end_page: int


@dataclass
class SlideIndexItem:
    ppt: Path
    slide_no: int
    fund_name_exact_keys: set[str]
    fund_name_exact_values: list[str]
    other_text_key: str
    page_no: str


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


# =========================================================
# GENERAL HELPERS
# =========================================================
def natural_key(text: str):
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", str(text))
    ]


def normalize_text(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "").strip())


def normalize_key(text: Any) -> str:
    return normalize_text(text).casefold()


def normalize_fund_search_text(text: Any) -> str:
    text = normalize_text(text)

    while text and not text[-1].isalnum():
        text = text[:-1].rstrip()

    return text


def normalize_fund_match_key(text: Any) -> str:
    """
    Normalize search text for matching by ignoring all special characters,
    symbols, punctuation, underscores, configured ignored characters like ß/¥,
    and extra spaces.

    Examples:
        Bandhan Gilt Fundßß      -> bandhan gilt fund
        Bandhan Gilt Fund ¥      -> bandhan gilt fund
        Bandhan-Gilt_Fund!!!     -> bandhan gilt fund
        Commentary - Equity      -> commentary equity
    """
    text = normalize_fund_search_text(text)
    cleaned_chars = []

    for ch in text:
        if ch in IGNORED_OBJECT_NAME_CHARS:
            cleaned_chars.append(" ")
            continue

        category = unicodedata.category(ch)

        if ch.isspace() or category.startswith("L") or category.startswith("N"):
            cleaned_chars.append(ch)
        else:
            cleaned_chars.append(" ")

    return normalize_text("".join(cleaned_chars)).casefold()


def index_search_terms(search_term: Any) -> list[str]:
    """
    Return all search terms to try for an IndexTable row.

    Example:
        Index row 'Snapshots' searches for 'Snapshots' and 'Equity Snapshot'.
    """
    terms = [normalize_fund_search_text(search_term)]
    search_key = normalize_fund_match_key(search_term)

    for alias_from, alias_terms in INDEX_SEARCH_ALIASES.items():
        if normalize_fund_match_key(alias_from) == search_key:
            terms.extend(alias_terms)

    deduped = []
    seen = set()

    for term in terms:
        term = normalize_fund_search_text(term)
        key = normalize_fund_match_key(term)

        if key and key not in seen:
            seen.add(key)
            deduped.append(term)

    return deduped


def normalized_set(items):
    return {normalize_key(x) for x in items}


def is_blacklisted(text: Any) -> bool:
    return normalize_key(text) in normalized_set(BLACKLIST_NAMES)


def is_pptx(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix.casefold() in PPT_EXTENSIONS
        and not path.name.startswith("~$")
    )


def iter_pptx_files(folder: Path):
    return sorted(
        (p for p in Path(folder).rglob("*") if is_pptx(p)),
        key=lambda p: natural_key(str(p.relative_to(folder))),
    )


def slide_count(pptx_path: Path) -> int:
    return len(Presentation(str(pptx_path)).slides)


def filename_contains(path: Path, keyword: str) -> bool:
    return keyword.casefold() in path.name.casefold()


def format_page_number(num: Any) -> str:
    num = str(num or "").strip()
    return f"0{num}" if num.isdigit() and len(num) == 1 else num


def first_number(text: Any) -> str | None:
    match = re.search(r"\d+", str(text or ""))
    return format_page_number(match.group(0)) if match else None


def save_replace_original(prs: Presentation, ppt_path: Path) -> None:
    ppt_path = Path(ppt_path)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        temp_path = Path(tmp.name)

    try:
        prs.save(str(temp_path))
        os.replace(temp_path, ppt_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


# =========================================================
# POWERPOINT SHAPE HELPERS
# =========================================================
def iter_shapes(shapes):
    for shape in shapes:
        yield shape

        if hasattr(shape, "shapes"):
            yield from iter_shapes(shape.shapes)


def normalize_object_name_for_match(text: Any) -> str:
    text = normalize_text(text)
    cleaned_chars = []

    for ch in text:
        if ch in IGNORED_OBJECT_NAME_CHARS:
            cleaned_chars.append(" ")
            continue

        category = unicodedata.category(ch)

        if ch.isspace() or category.startswith("L") or category.startswith("N"):
            cleaned_chars.append(ch)
        else:
            cleaned_chars.append(" ")

    return normalize_text("".join(cleaned_chars)).casefold()


def normalized_object_name_set(items):
    return {normalize_object_name_for_match(x) for x in items}


OBJECTS_TO_SEARCH_IN_KEY = normalized_object_name_set(OBJECTS_TO_SEARCH_IN)


def get_shape_name_key(shape) -> str:
    return normalize_object_name_for_match(getattr(shape, "name", ""))


def shape_text(shape) -> str:
    parts = []

    if getattr(shape, "has_text_frame", False):
        parts.append(shape.text or "")

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text or "")

    return "\n".join(x for x in parts if x)


def table_shapes(slide):
    for shape in iter_shapes(slide.shapes):
        if getattr(shape, "has_table", False):
            yield shape


def find_named_table_shapes(slide, table_object_name: str):
    table_object_key = normalize_object_name_for_match(table_object_name)

    for shape in table_shapes(slide):
        if normalize_object_name_for_match(getattr(shape, "name", "")) == table_object_key:
            yield shape


# =========================================================
# COLOR HELPERS
# =========================================================
def rgb_to_hex(rgb) -> str:
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def hex_to_rgbcolor(hex_color: str) -> RGBColor:
    hex_color = hex_color.strip().replace("#", "")

    if len(hex_color) != 6:
        raise ValueError("Hex color must be in #RRGGBB format.")

    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


def color_matches(rgb, expected) -> bool:
    return rgb[0] == expected[0] and rgb[1] == expected[1] and rgb[2] == expected[2]


def get_color_mode_from_rgb(rgb):
    if color_matches(rgb, FONT_COLOR):
        return "Dark", rgb_to_hex(FONT_COLOR)

    if color_matches(rgb, FONT_COLOR_LIGHT):
        return "Light", rgb_to_hex(FONT_COLOR_LIGHT)

    return "Custom", rgb_to_hex(rgb)


def get_rgb_from_mode(mode: str, custom_hex: str | None = None) -> RGBColor:
    if mode == "Dark":
        return FONT_COLOR

    if mode == "Light":
        return FONT_COLOR_LIGHT

    if mode == "Custom" and custom_hex:
        return hex_to_rgbcolor(custom_hex)

    return FONT_COLOR


def get_font_name_from_mode(mode: str) -> str:
    return FONT_NAME_LIGHT if mode == "Light" else FONT_NAME_DARK


def get_run_rgb(run):
    try:
        rgb = run.font.color.rgb
        return rgb if rgb else None
    except Exception:
        return None


# =========================================================
# PAGE NUMBER HELPERS
# =========================================================
def normalize_shape_name(name: Any) -> str:
    return " ".join(str(name or "").strip().split())


def normalize_shape_key(name: Any) -> str:
    return normalize_shape_name(name).replace(" ", "").casefold()


PAGE_NUMBER_ALIAS_KEYS = {normalize_shape_key(n) for n in SHAPE_NAME_ALIASES}


def is_page_number_shape(shape) -> bool:
    key = normalize_shape_key(getattr(shape, "name", ""))

    return any(
        key == alias_key or key.startswith(alias_key)
        for alias_key in PAGE_NUMBER_ALIAS_KEYS
    )


def is_landscape(prs: Presentation) -> bool:
    return prs.slide_width >= prs.slide_height


def get_default_pos(prs: Presentation):
    return LANDSCAPE_POS if is_landscape(prs) else PORTRAIT_POS


def find_page_number_shapes(slide) -> list:
    if slide is None:
        return []

    return [
        shape
        for shape in iter_shapes(slide.shapes)
        if is_page_number_shape(shape)
    ]


def find_first_page_number_shape(slide):
    matches = find_page_number_shapes(slide)
    return matches[0] if matches else None


def duplicate_shape_to_slide(target_slide, source_shape):
    new_el = deepcopy(source_shape._element)

    cNvPr_items = new_el.xpath(".//p:cNvPr")

    if cNvPr_items:
        cNvPr_items[0].set("id", str(target_slide.shapes._next_shape_id))

    target_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
    return target_slide.shapes[-1]


def set_text_keep_format(shape, text: Any, color: RGBColor, font_name: str) -> bool:
    if not getattr(shape, "has_text_frame", False):
        return False

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = False
    tf.vertical_anchor = VERTICAL_ALIGNMENT
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    paragraph = tf.paragraphs[0]
    paragraph.alignment = PARAGRAPH_ALIGNMENT

    run = paragraph.add_run()
    run.text = str(text)
    run.font.name = font_name
    run.font.size = FONT_SIZE
    run.font.bold = FONT_BOLD
    run.font.italic = FONT_ITALIC
    run.font.color.rgb = color

    return True


def create_page_number_box(slide, prs: Presentation, color: RGBColor, font_name: str):
    pos = get_default_pos(prs)

    shape = slide.shapes.add_textbox(
        pos["left"],
        pos["top"],
        pos["width"],
        pos["height"],
    )

    shape.name = SHAPE_NAME
    set_text_keep_format(shape, "", color, font_name)

    return shape


def prepare_page_number_shape(slide, prs: Presentation, first_slide_template, color: RGBColor, font_name: str):
    existing_shapes = find_page_number_shapes(slide)

    if existing_shapes:
        for shape in existing_shapes:
            shape.name = SHAPE_NAME
        return existing_shapes[0]

    if first_slide_template:
        copied = duplicate_shape_to_slide(slide, first_slide_template)
        copied.name = SHAPE_NAME
        return copied

    return create_page_number_box(slide, prs, color, font_name)


def set_page_number_on_slide(
    slide,
    prs: Presentation,
    page_no: Any,
    color: RGBColor,
    font_name: str,
    first_slide_template=None,
) -> int:
    page_shapes = find_page_number_shapes(slide)

    if not page_shapes:
        page_shapes = [
            prepare_page_number_shape(
                slide=slide,
                prs=prs,
                first_slide_template=first_slide_template,
                color=color,
                font_name=font_name,
            )
        ]

    updated = 0

    for shape in page_shapes:
        shape.name = SHAPE_NAME

        if set_text_keep_format(shape, page_no, color, font_name):
            updated += 1

    return updated


def get_page_number_from_slide(slide) -> str | None:
    for shape in iter_shapes(slide.shapes):
        if is_page_number_shape(shape):
            num = first_number(shape_text(shape))

            if num:
                return num

    return None


def detect_page_number_color_mode(pptx_path: Path):
    try:
        prs = Presentation(str(pptx_path))

        for slide in prs.slides:
            for shape in find_page_number_shapes(slide):
                if not getattr(shape, "has_text_frame", False):
                    continue

                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        rgb = get_run_rgb(run)

                        if rgb:
                            return get_color_mode_from_rgb(rgb)

    except Exception:
        pass

    return "Dark", rgb_to_hex(FONT_COLOR)


# =========================================================
# TEXT REPLACEMENT HELPERS
# =========================================================
def get_text_frame_text(text_frame) -> str:
    return "\n".join(p.text for p in text_frame.paragraphs)


def copy_font_style(src_font, dest_font) -> None:
    if not src_font:
        return

    dest_font.name = src_font.name
    dest_font.size = src_font.size
    dest_font.bold = src_font.bold
    dest_font.italic = src_font.italic
    dest_font.underline = src_font.underline

    try:
        if src_font.color and src_font.color.rgb:
            dest_font.color.rgb = src_font.color.rgb
    except Exception:
        pass


def replace_text_in_text_frame_keep_format(text_frame, new_text: Any) -> None:
    font_template = None
    para_alignment = None

    for paragraph in text_frame.paragraphs:
        if para_alignment is None:
            para_alignment = paragraph.alignment

        for run in paragraph.runs:
            font_template = run.font
            break

        if font_template:
            break

    text_frame.clear()

    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = para_alignment if para_alignment is not None else PP_ALIGN.RIGHT

    run = paragraph.add_run()
    run.text = str(new_text)

    if font_template:
        copy_font_style(font_template, run.font)
    else:
        run.font.name = DEFAULT_FONT_NAME
        run.font.size = Pt(8)


def apply_replacements_to_shape(shape, replacements: dict[str, str]) -> int:
    if not getattr(shape, "has_text_frame", False):
        return 0

    original = get_text_frame_text(shape.text_frame)
    updated = original

    for old, new in replacements.items():
        updated = updated.replace(old, new)

    if updated != original:
        replace_text_in_text_frame_keep_format(shape.text_frame, updated)
        return 1

    return 0


def apply_replacements_to_table(table, replacements: dict[str, str]) -> int:
    updated_cells = 0

    for row in table.rows:
        for cell in row.cells:
            original = "\n".join(p.text for p in cell.text_frame.paragraphs)
            updated = original

            for old, new in replacements.items():
                updated = updated.replace(old, new)

            if updated != original:
                replace_text_in_text_frame_keep_format(cell.text_frame, updated)
                updated_cells += 1

    return updated_cells


def build_replacements_for_file(
    file_meta: FileMeta,
    all_files_map: dict[str, FileMeta],
    total_pages: int,
    slide_index: int,
) -> dict[str, str]:
    page_no = file_meta.start_page + slide_index - 1

    replacements = {
        "{{PAGE_NO}}": str(page_no),
        "{{START_PAGE}}": str(file_meta.start_page),
        "{{TOTAL_PAGES}}": str(total_pages),
        "{{PAGE_X_OF_Y}}": f"Page {page_no} of {total_pages}",
    }

    for key, meta in all_files_map.items():
        replacements[f"{{{{{key}}}}}"] = str(meta.start_page)
        replacements[f"{{{{REF:{key}}}}}"] = str(meta.start_page)

    return replacements


# =========================================================
# INDEX SEARCH HELPERS
# =========================================================
def add_exact_search_candidate(candidates: list[str], seen_keys: set[str], text: Any) -> None:
    cleaned = normalize_fund_search_text(text)
    key = normalize_fund_match_key(cleaned)

    if key and key not in seen_keys:
        seen_keys.add(key)
        candidates.append(cleaned)


def shape_search_candidates(shape) -> list[str]:
    candidates = []
    seen_keys = set()

    if getattr(shape, "has_text_frame", False):
        add_exact_search_candidate(candidates, seen_keys, shape.text or "")

        try:
            for paragraph in shape.text_frame.paragraphs:
                add_exact_search_candidate(candidates, seen_keys, paragraph.text)
        except Exception:
            pass

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                add_exact_search_candidate(candidates, seen_keys, cell.text or "")

                try:
                    for paragraph in cell.text_frame.paragraphs:
                        add_exact_search_candidate(candidates, seen_keys, paragraph.text)
                except Exception:
                    pass

    return candidates


def get_searchable_exact_candidates_from_slide(slide) -> list[str]:
    candidates = []
    seen_keys = set()
    fund_name_key = normalize_object_name_for_match("Fund Name")

    for shape in iter_shapes(slide.shapes):
        if get_shape_name_key(shape) == fund_name_key:
            for candidate in shape_search_candidates(shape):
                key = normalize_fund_match_key(candidate)

                if key and key not in seen_keys:
                    seen_keys.add(key)
                    candidates.append(candidate)

    return candidates


def get_other_searchable_text_from_slide(slide) -> str:
    parts = []
    fund_name_key = normalize_object_name_for_match("Fund Name")

    for shape in iter_shapes(slide.shapes):
        shape_key = get_shape_name_key(shape)

        if shape_key in OBJECTS_TO_SEARCH_IN_KEY and shape_key != fund_name_key:
            parts.append(shape_text(shape))

    return "\n".join(x for x in parts if x)


def build_slide_index(search_folder: Path, log: Callable[[str], None] | None = None) -> list[SlideIndexItem]:
    index = []

    for ppt_path in iter_pptx_files(search_folder):
        try:
            prs = Presentation(str(ppt_path))
        except Exception as exc:
            if log:
                log(f"[Tab7] WARN: Could not open for search: {ppt_path.name} | {exc}")
            continue

        for slide_no, slide in enumerate(prs.slides, start=1):
            fund_candidates = get_searchable_exact_candidates_from_slide(slide)
            fund_keys = {
                normalize_fund_match_key(candidate)
                for candidate in fund_candidates
                if normalize_fund_match_key(candidate)
            }

            other_text = normalize_text(get_other_searchable_text_from_slide(slide))
            other_text_key = normalize_fund_match_key(other_text)

            page_no = get_page_number_from_slide(slide)

            if page_no and (fund_keys or other_text_key):
                index.append(
                    SlideIndexItem(
                        ppt=ppt_path.resolve(),
                        slide_no=slide_no,
                        fund_name_exact_keys=fund_keys,
                        fund_name_exact_values=fund_candidates,
                        other_text_key=other_text_key,
                        page_no=page_no,
                    )
                )

    return index


def find_first_page_number(
    search_index: list[SlideIndexItem],
    search_term: Any,
    current_ppt_path: Path | None = None,
) -> str | None:
    current_ppt_path = Path(current_ppt_path).resolve() if current_ppt_path else None

    for candidate_term in index_search_terms(search_term):
        term_key = normalize_fund_match_key(candidate_term)
        fallback_term_key = normalize_key(normalize_fund_search_text(candidate_term))

        if not term_key and not fallback_term_key:
            continue

        for item in search_index:
            if EXCLUDE_CURRENT_UPDATE_PPT and current_ppt_path and item.ppt == current_ppt_path:
                continue

            # Fund Name exact matching ignores symbols/special characters.
            if term_key and term_key in item.fund_name_exact_keys:
                return item.page_no

            # Other configured objects use substring matching, also ignoring symbols.
            if term_key and term_key in item.other_text_key:
                return item.page_no

            # Fallback keeps previous raw substring behavior.
            if fallback_term_key and fallback_term_key in item.other_text_key:
                return item.page_no

    return None


def apply_default_font_if_missing(run) -> None:
    if run.font.name is None:
        run.font.name = DEFAULT_FONT_NAME

    if run.font.size is None:
        run.font.size = DEFAULT_FONT_SIZE


def set_cell_text_keep_format(cell, new_text: Any) -> None:
    new_text = "" if new_text is None else str(new_text)
    tf = cell.text_frame

    if not tf.paragraphs:
        cell.text = new_text
        paragraph = cell.text_frame.paragraphs[0]
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        apply_default_font_if_missing(run)
        return

    first_para = tf.paragraphs[0]

    if not first_para.runs:
        run = first_para.add_run()
        run.text = new_text
    else:
        run = first_para.runs[0]
        run.text = new_text

        for extra_run in first_para.runs[1:]:
            extra_run.text = ""

    apply_default_font_if_missing(run)

    for paragraph in tf.paragraphs[1:]:
        for extra_run in paragraph.runs:
            extra_run.text = ""


def update_index_page_ppt(
    ppt_path: Path,
    search_index: list[SlideIndexItem],
    log: Callable[[str], None] | None = None,
) -> int:
    prs = Presentation(str(ppt_path))
    changes = 0

    for slide_no, slide in enumerate(prs.slides, start=1):
        for shape in find_named_table_shapes(slide, INDEX_TABLE_OBJECT_NAME):
            table = shape.table

            if len(table.columns) <= INDEX_PAGE_NO_COL:
                continue

            for row_no, row in enumerate(table.rows, start=1):
                original_search_text = normalize_text(row.cells[INDEX_SEARCH_COL].text)
                search_text = normalize_fund_search_text(original_search_text)

                if not search_text or is_blacklisted(search_text):
                    continue

                page_no = find_first_page_number(
                    search_index,
                    search_text,
                    current_ppt_path=ppt_path,
                )

                if page_no:
                    set_cell_text_keep_format(row.cells[INDEX_PAGE_NO_COL], page_no)
                    changes += 1

                    if log:
                        log(
                            f"[Tab7] [IndexTable] {ppt_path.name} | "
                            f"Slide {slide_no} | Row {row_no}: "
                            f"'{original_search_text}' -> {page_no}"
                        )

    if changes:
        save_replace_original(prs, ppt_path)

    return changes


def update_fundmanager_wise_ppt(
    ppt_path: Path,
    search_index: list[SlideIndexItem],
    log: Callable[[str], None] | None = None,
) -> int:
    prs = Presentation(str(ppt_path))
    changes = 0

    for slide_no, slide in enumerate(prs.slides, start=1):
        for shape in table_shapes(slide):
            table = shape.table

            if len(table.columns) <= FUNDMANAGER_PAGE_NO_COL:
                continue

            for row_no, row in enumerate(table.rows, start=1):
                original_search_text = normalize_text(row.cells[FUNDMANAGER_SEARCH_COL].text)
                search_text = normalize_fund_search_text(original_search_text)

                if not search_text or is_blacklisted(search_text):
                    continue

                page_no = find_first_page_number(
                    search_index,
                    search_text,
                    current_ppt_path=ppt_path,
                )

                if page_no:
                    set_cell_text_keep_format(row.cells[FUNDMANAGER_PAGE_NO_COL], page_no)
                    changes += 1

                    if log:
                        log(
                            f"[Tab7] [FundManager_Wise] {ppt_path.name} | "
                            f"Slide {slide_no} | Row {row_no}: "
                            f"'{original_search_text}' -> {page_no}"
                        )

    if changes:
        save_replace_original(prs, ppt_path)

    return changes


def update_index_and_fundmanager_references(search_folder: Path, progress_cb=None) -> int:
    def log(msg: str):
        if progress_cb:
            progress_cb(msg)

    log("[Tab7] Building search index from updated output folder...")
    log(f"[Tab7] Searching only in objects: {OBJECTS_TO_SEARCH_IN}")
    log(f"[Tab7] Index table object name: {INDEX_TABLE_OBJECT_NAME}")
    log("[Tab7] Fund Name object uses exact full-text matching.")
    log("[Tab7] Other configured objects use substring matching.")

    search_index = build_slide_index(search_folder, log=log)
    log(f"[Tab7] Search index built. Slides indexed: {len(search_index)}")

    total_changes = 0

    for ppt_path in iter_pptx_files(search_folder):
        try:
            if filename_contains(ppt_path, "Index page"):
                total_changes += update_index_page_ppt(ppt_path, search_index, log=log)

            elif filename_contains(ppt_path, "FundManager_Wise"):
                total_changes += update_fundmanager_wise_ppt(ppt_path, search_index, log=log)

        except Exception as exc:
            log(f"[Tab7] ERROR: Failed updating references in {ppt_path.name} | {exc}")

    log(f"[Tab7] Reference update completed. Total cells updated: {total_changes}")

    return total_changes


# =========================================================
# MAIN PROCESSING
# =========================================================
def update_references_and_page_numbers(
    rows: list[dict],
    output_folder: Path,
    progress_cb=None,
) -> Path:
    def log(msg: str):
        if progress_cb:
            progress_cb(msg)

    if not rows:
        raise ValueError("No PPTX rows provided.")

    output_folder.mkdir(parents=True, exist_ok=True)

    files_map: dict[str, FileMeta] = {}
    total_pages = 0

    for row in rows:
        pptx_path = Path(row["path"])
        slides = int(row["slides"])
        start_page = int(row["start_page"])
        end_page = start_page + slides - 1
        total_pages = max(total_pages, end_page)

        files_map[pptx_path.stem] = FileMeta(
            file_name=pptx_path.name,
            file_stem=pptx_path.stem,
            path=str(pptx_path),
            slides=slides,
            start_page=start_page,
            end_page=end_page,
        )

        log(
            f"[Tab7] {pptx_path.name}: slides={slides}, "
            f"start={start_page}, end={end_page}, color={row['color_mode']}"
        )

    total_slides = sum(int(row["slides"]) for row in rows)
    processed_slides = 0
    last_percent = -1

    log(f"[Tab7] Total pages across selected PPTX files: {total_pages}")
    log("[Tab7] Progress: 0%")

    for row in rows:
        pptx_path = Path(row["path"])
        file_meta = files_map[pptx_path.stem]

        color_mode = row["color_mode"]
        custom_color = row["custom_color"]

        color = get_rgb_from_mode(color_mode, custom_color)
        font_name = get_font_name_from_mode(color_mode)

        prs = Presentation(str(pptx_path))

        first_slide = prs.slides[0] if prs.slides else None
        first_slide_template = find_first_page_number_shape(first_slide)

        total_replacements = 0
        total_page_shapes_updated = 0
        page_no = file_meta.start_page

        for slide_idx, slide in enumerate(prs.slides, start=1):
            replacements = build_replacements_for_file(
                file_meta=file_meta,
                all_files_map=files_map,
                total_pages=total_pages,
                slide_index=slide_idx,
            )

            for shape in iter_shapes(slide.shapes):
                if is_page_number_shape(shape):
                    continue

                if getattr(shape, "has_table", False):
                    total_replacements += apply_replacements_to_table(shape.table, replacements)
                else:
                    total_replacements += apply_replacements_to_shape(shape, replacements)

            found_before = len(find_page_number_shapes(slide))

            updated_count = set_page_number_on_slide(
                slide=slide,
                prs=prs,
                page_no=page_no,
                color=color,
                font_name=font_name,
                first_slide_template=first_slide_template,
            )

            found_after = len(find_page_number_shapes(slide))
            total_page_shapes_updated += updated_count

            log(
                f"[Tab7] {pptx_path.name} | Slide {slide_idx}: "
                f"page={page_no}, found_before={found_before}, "
                f"found_after={found_after}, updated={updated_count}"
            )

            page_no += 1
            processed_slides += 1

            if total_slides > 0:
                percent = int((processed_slides / total_slides) * 80)

                if percent != last_percent:
                    log(f"[Tab7] Progress: {percent}%")
                    last_percent = percent

        output_path = output_folder / pptx_path.name

        if output_path.resolve() == pptx_path.resolve():
            save_replace_original(prs, output_path)
        else:
            prs.save(str(output_path))

        log(
            f"[Tab7] Updated {pptx_path.name} -> "
            f"token replacements: {total_replacements}, "
            f"page-number shapes updated: {total_page_shapes_updated}, "
            f"color mode: {color_mode}, font: {font_name}"
        )

    log("[Tab7] Page numbers added. Starting Index/FundManager reference update...")
    log("[Tab7] Progress: 85%")

    update_index_and_fundmanager_references(
        search_folder=output_folder,
        progress_cb=progress_cb,
    )

    log("[Tab7] Progress: 100%")

    return output_folder


# =========================================================
# ROW SELECTION BORDER
# =========================================================
class RowBorderDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        selected = bool(opt.state & QStyle.StateFlag.State_Selected)

        if selected:
            opt.state &= ~QStyle.StateFlag.State_Selected
            opt.state &= ~QStyle.StateFlag.State_HasFocus

        super().paint(painter, opt, index)

        if not selected or index.column() != 0:
            return

        table = self.parent()
        row = index.row()

        left = table.visualRect(table.model().index(row, 0)).left()
        right = table.visualRect(table.model().index(row, table.columnCount() - 1)).right()

        rect = QRect(
            left,
            option.rect.top(),
            right - left,
            option.rect.height() - 1,
        )

        painter.save()
        painter.setRenderHint(painter.RenderHint.Antialiasing, False)
        painter.setPen(QPen(QColor("#2563EB"), 2))
        painter.drawRect(rect.adjusted(1, 1, -1, -2))
        painter.restore()


# =========================================================
# TABLE WIDGET
# =========================================================
class PptTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(0, 5, parent)

        self._press_pos = None
        self._source_row = -1
        self._is_dragging_row = False

        self.setItemDelegate(RowBorderDelegate(self))

        self.setHorizontalHeaderLabels(
            [
                "File Name",
                "Slides",
                "Start Page Number",
                "Color Mode",
                "Custom Color",
            ]
        )

        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        self.setDragEnabled(False)
        self.setAcceptDrops(False)
        self.viewport().setAcceptDrops(False)
        self.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)

        self.setSortingEnabled(False)
        self.verticalHeader().setVisible(False)

        header = self.horizontalHeader()
        header.setMinimumSectionSize(60)
        header.setStretchLastSection(True)

        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        self.setColumnWidth(0, 430)
        self.setColumnWidth(1, 80)
        self.setColumnWidth(2, 150)
        self.setColumnWidth(3, 130)
        self.setColumnWidth(4, 150)

        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)

        self.setShowGrid(True)
        self.setAlternatingRowColors(False)

        self.setStyleSheet(
            """
            QTableWidget::item {
                padding: 4px;
            }

            QTableWidget::item:selected {
                background: transparent;
                color: inherit;
            }

            QHeaderView::section {
                padding: 4px;
            }
            """
        )

        self.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )

    def resizeEvent(self, event):
        super().resizeEvent(event)

        viewport_width = max(0, self.viewport().width())

        if viewport_width <= 0:
            return

        slides_w = 80
        start_w = 150
        mode_w = 130
        color_w = 150

        remaining_for_file = viewport_width - slides_w - start_w - mode_w - color_w - 30
        file_w = max(260, min(520, remaining_for_file))

        if viewport_width >= 850:
            self.setColumnWidth(0, file_w)
            self.setColumnWidth(1, slides_w)
            self.setColumnWidth(2, start_w)
            self.setColumnWidth(3, mode_w)
            self.setColumnWidth(4, color_w)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._press_pos = event.position().toPoint()
            self._source_row = self.rowAt(self._press_pos.y())
            self._is_dragging_row = False

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not event.buttons() & Qt.MouseButton.LeftButton:
            return super().mouseMoveEvent(event)

        if self._source_row < 0 or self._press_pos is None:
            return super().mouseMoveEvent(event)

        distance = (event.position().toPoint() - self._press_pos).manhattanLength()

        if distance >= QApplication.startDragDistance():
            self._is_dragging_row = True
            self.setCursor(Qt.CursorShape.ClosedHandCursor)
            self.selectRow(self._source_row)
            event.accept()
            return

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if (
            event.button() == Qt.MouseButton.LeftButton
            and self._is_dragging_row
            and self._source_row >= 0
        ):
            self.unsetCursor()

            drop_pos = event.position().toPoint()
            target_row = self.rowAt(drop_pos.y())

            if target_row < 0:
                target_row = self.rowCount()
            else:
                item = self.item(target_row, 0)

                if item:
                    rect = self.visualItemRect(item)

                    if drop_pos.y() > rect.center().y():
                        target_row += 1

            self.move_row(self._source_row, target_row)

            self._press_pos = None
            self._source_row = -1
            self._is_dragging_row = False

            event.accept()
            return

        self.unsetCursor()
        self._press_pos = None
        self._source_row = -1
        self._is_dragging_row = False

        super().mouseReleaseEvent(event)

    def move_row(self, source_row: int, target_row: int):
        if source_row < 0 or source_row >= self.rowCount():
            return

        target_row = max(0, min(target_row, self.rowCount()))

        if source_row == target_row or source_row + 1 == target_row:
            return

        parent = self.parent()

        if parent and hasattr(parent, "updating_table"):
            parent.updating_table = True

        try:
            row_items = []

            for col in range(self.columnCount()):
                item = self.item(source_row, col)
                row_items.append(item.clone() if item else QTableWidgetItem(""))

            self.removeRow(source_row)

            if target_row > source_row:
                target_row -= 1

            self.insertRow(target_row)

            for col, item in enumerate(row_items):
                self.setItem(target_row, col, item)

            self.selectRow(target_row)

        finally:
            if parent and hasattr(parent, "updating_table"):
                parent.updating_table = False

        if parent and hasattr(parent, "auto_fill_from_row"):
            parent.auto_fill_from_row(0)


# =========================================================
# MAIN UI
# =========================================================
class UpdateRefsWidget(QWidget):
    def __init__(self, state: Any | None = None):
        super().__init__()

        self.state = state
        self.thread_pool = QThreadPool.globalInstance()
        self.updating_table = False

        self.click_order_mode = False
        self.click_order_sequence = []
        self.default_output_folder = self._resolve_default_output_folder()

        self._build_ui()
        self._apply_default_output_folder()
        self._set_progress_idle()

    def _app_run_folder(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent

        return Path(__file__).resolve().parent

    def _resolve_default_output_folder(self) -> Path:
        state_folder = getattr(self.state, "tab7_output_folder", None) if self.state is not None else None

        if state_folder:
            folder = Path(state_folder)
        else:
            output_root = getattr(self.state, "output_root", None) if self.state is not None else None
            folder = (Path(output_root) if output_root else self._app_run_folder() / "output") / "07_ppt_page_number"

        folder.mkdir(parents=True, exist_ok=True)

        if self.state is not None:
            setattr(self.state, "tab7_output_folder", folder)

        return folder

    def _apply_default_output_folder(self) -> None:
        folder = self._resolve_default_output_folder()
        self.default_output_folder = folder

        if not self.output_folder_edit.text().strip():
            self.output_folder_edit.setText(str(folder))

    def _build_ui(self):
        self.setWindowTitle(APP_TITLE)

        root = QHBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(8)

        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)

        root.addWidget(splitter)

        left_widget = QWidget()
        left_panel = QVBoxLayout(left_widget)
        left_panel.setSpacing(8)
        left_panel.setContentsMargins(0, 0, 0, 0)

        title = QLabel(APP_TITLE)
        title.setStyleSheet("font-weight: bold; font-size: 14px;")
        left_panel.addWidget(title)

        help_label = QLabel(
            "Professional PPTX update tool.\n"
            "- Index page table object must be named IndexTable\n"
            "- IndexTable Column 1 is search text\n"
            "- IndexTable Column 2 is page number output\n"
            "- PageNumber/Page Number shapes are updated or created\n"
            "- Object-name matching ignores ß, ¥, punctuation, symbols, and underscores\n"
            "- Fund Name uses exact full-text matching\n"
            "- Other configured objects use substring matching"
        )
        help_label.setWordWrap(True)
        left_panel.addWidget(help_label)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Update logs will appear here...")
        left_panel.addWidget(self.log_box, stretch=1)

        self.btn_clear_log = QPushButton("Clear Log")
        left_panel.addWidget(self.btn_clear_log)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)

        right_container = QWidget()
        right_panel = QVBoxLayout(right_container)
        right_panel.setSpacing(8)
        right_panel.setContentsMargins(0, 0, 0, 0)

        group = QGroupBox("Inputs")
        grid = QGridLayout(group)

        self.input_folder_edit = QLineEdit()
        self.btn_browse_input = QPushButton("Browse...")

        grid.addWidget(QLabel("Input PPTX Folder:"), 0, 0)
        grid.addWidget(self.input_folder_edit, 0, 1)
        grid.addWidget(self.btn_browse_input, 0, 2)

        self.output_folder_edit = QLineEdit()
        self.btn_browse_output = QPushButton("Browse...")

        grid.addWidget(QLabel("Output Folder:"), 1, 0)
        grid.addWidget(self.output_folder_edit, 1, 1)
        grid.addWidget(self.btn_browse_output, 1, 2)

        self.start_page_edit = QLineEdit("1")

        grid.addWidget(QLabel("Initial Start Page:"), 2, 0)
        grid.addWidget(self.start_page_edit, 2, 1)

        right_panel.addWidget(group)

        table_group = QGroupBox("Files / Order / Start Pages / Color Mode")
        table_layout = QVBoxLayout(table_group)

        self.table = PptTable(self)
        self.table.setMinimumHeight(360)
        self.table.setSizeAdjustPolicy(QAbstractItemView.SizeAdjustPolicy.AdjustIgnored)

        self.table.itemChanged.connect(self.on_item_changed)
        self.table.cellDoubleClicked.connect(self.choose_color)
        self.table.cellClicked.connect(self._handle_click_order_cell)

        table_layout.addWidget(self.table, stretch=1)

        table_btn_row = QHBoxLayout()

        self.btn_load_files = QPushButton("Load Files From Input Folder")
        self.btn_auto_fill = QPushButton("Auto Fill Continuous Start Pages")
        self.btn_click_order = QPushButton("Click Order Mode")
        self.btn_up = QPushButton("Move Up")
        self.btn_down = QPushButton("Move Down")

        table_btn_row.addWidget(self.btn_load_files)
        table_btn_row.addWidget(self.btn_auto_fill)
        table_btn_row.addWidget(self.btn_click_order)
        table_btn_row.addWidget(self.btn_up)
        table_btn_row.addWidget(self.btn_down)
        table_btn_row.addStretch(1)

        table_layout.addLayout(table_btn_row)
        right_panel.addWidget(table_group, stretch=1)

        progress_group = QGroupBox("Update Progress")
        progress_layout = QVBoxLayout(progress_group)

        self.progress_status_label = QLabel("Idle")
        progress_layout.addWidget(self.progress_status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setTextVisible(True)
        progress_layout.addWidget(self.progress_bar)

        right_panel.addWidget(progress_group)

        run_row = QHBoxLayout()
        run_row.addStretch(1)

        self.btn_run = QPushButton("Update References")
        run_row.addWidget(self.btn_run)

        right_panel.addLayout(run_row)

        right_scroll.setWidget(right_container)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([360, 980])

        self.btn_browse_input.clicked.connect(self._browse_input)
        self.btn_browse_output.clicked.connect(self._browse_output)
        self.btn_load_files.clicked.connect(self._load_file_table)
        self.btn_auto_fill.clicked.connect(lambda: self.auto_fill_from_row(0))
        self.btn_click_order.clicked.connect(self._toggle_click_order_mode)
        self.btn_up.clicked.connect(self._move_up)
        self.btn_down.clicked.connect(self._move_down)
        self.btn_clear_log.clicked.connect(self.log_box.clear)
        self.btn_run.clicked.connect(self._run_update)

    def _browse_input(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select input folder",
            self.input_folder_edit.text().strip() or str(Path.cwd()),
        )

        if folder:
            self.input_folder_edit.setText(folder)

            if not self.output_folder_edit.text().strip():
                self.output_folder_edit.setText(str(self._resolve_default_output_folder()))

            self._load_file_table()

    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select output folder",
            self.output_folder_edit.text().strip() or str(self._resolve_default_output_folder()),
        )

        if folder:
            self.output_folder_edit.setText(folder)

    def _log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

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

    def _extract_progress_value(self, payload):
        if isinstance(payload, (int, float)):
            return max(0, min(100, int(payload)))

        if isinstance(payload, str):
            text = payload.strip()

            if "Progress:" not in text and not text.endswith("%"):
                return None

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
            self._log(f"[Tab7] Progress update: {payload}")

    def apply_color_preview(self, item, hex_color: str):
        item.setBackground(QColor(hex_color))

        if hex_color.upper() in ["#FFFFFF", "#FFFFFE", "#FDFDFD", "#FEFEFE"]:
            item.setForeground(QColor("#000000"))
        else:
            item.setForeground(QColor("#FFFFFF"))

    def _load_file_table(self):
        self.click_order_mode = False
        self.click_order_sequence = []
        self.btn_click_order.setText("Click Order Mode")
        self.table.setRowCount(0)

        input_folder_text = self.input_folder_edit.text().strip()

        if not input_folder_text:
            return

        input_folder = Path(input_folder_text)

        if not input_folder.exists() or not input_folder.is_dir():
            return

        try:
            initial_start = int(self.start_page_edit.text().strip())

            if initial_start < 1:
                initial_start = 1

        except Exception:
            initial_start = 1

        self.updating_table = True
        page_no = initial_start

        files = sorted(input_folder.glob("*.pptx"), key=lambda p: natural_key(p.name))

        for pptx_path in files:
            if pptx_path.name.startswith("~$"):
                continue

            row = self.table.rowCount()
            self.table.insertRow(row)

            try:
                slides = slide_count(pptx_path)
            except Exception:
                slides = 0

            color_mode, custom_color = detect_page_number_color_mode(pptx_path)

            file_item = QTableWidgetItem(pptx_path.name)
            file_item.setData(Qt.ItemDataRole.UserRole, pptx_path.name)
            file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

            slides_item = QTableWidgetItem(str(slides))
            slides_item.setFlags(slides_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

            start_item = QTableWidgetItem(str(page_no))

            mode_item = QTableWidgetItem(color_mode)
            mode_item.setFlags(mode_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

            custom_item = QTableWidgetItem(custom_color)
            custom_item.setFlags(custom_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.apply_color_preview(custom_item, custom_color)

            self.table.setItem(row, 0, file_item)
            self.table.setItem(row, 1, slides_item)
            self.table.setItem(row, 2, start_item)
            self.table.setItem(row, 3, mode_item)
            self.table.setItem(row, 4, custom_item)

            page_no += slides

        self.updating_table = False

    def _base_filename(self, row: int) -> str:
        item = self.table.item(row, 0)

        if not item:
            return ""

        original = item.data(Qt.ItemDataRole.UserRole)
        return str(original or item.text()).strip()

    def _toggle_click_order_mode(self):
        if self.table.rowCount() == 0:
            QMessageBox.information(self, "No Files", "Load PPTX files first.")
            return

        self.click_order_mode = not self.click_order_mode

        if self.click_order_mode:
            self.click_order_sequence = []
            self.btn_click_order.setText("Cancel Click Order")
            self._refresh_click_order_marks()
            self._log("[Tab7] Click Order Mode enabled. Click files in required order.")
        else:
            self._clear_click_order_marks()
            self.btn_click_order.setText("Click Order Mode")
            self._log("[Tab7] Click Order Mode cancelled.")

    def _handle_click_order_cell(self, row: int, col: int):
        if not self.click_order_mode:
            return

        if row < 0 or row >= self.table.rowCount():
            return

        filename = self._base_filename(row)

        if row in self.click_order_sequence:
            self.click_order_sequence.remove(row)
            self._log(f"[Tab7] Removed from click order: {filename}")
        else:
            self.click_order_sequence.append(row)
            self._log(
                f"[Tab7] Click order selected "
                f"{len(self.click_order_sequence)}/{self.table.rowCount()}: {filename}"
            )

        self._refresh_click_order_marks()

        if len(self.click_order_sequence) == self.table.rowCount():
            self._apply_click_order()

    def _refresh_click_order_marks(self):
        self.updating_table = True

        try:
            order_map = {
                row: index + 1
                for index, row in enumerate(self.click_order_sequence)
            }

            for row in range(self.table.rowCount()):
                file_item = self.table.item(row, 0)

                if not file_item:
                    continue

                base_name = self._base_filename(row)
                file_item.setText(
                    f"{order_map[row]}. {base_name}" if row in order_map else base_name
                )

        finally:
            self.updating_table = False

    def _clear_click_order_marks(self):
        self.updating_table = True

        try:
            for row in range(self.table.rowCount()):
                file_item = self.table.item(row, 0)

                if file_item:
                    file_item.setText(self._base_filename(row))

            self.click_order_sequence = []

        finally:
            self.updating_table = False

    def _apply_click_order(self):
        self._log("[Tab7] Applying clicked file order...")

        self.updating_table = True

        try:
            for row in range(self.table.rowCount()):
                file_item = self.table.item(row, 0)

                if file_item:
                    file_item.setText(self._base_filename(row))

            ordered_rows = []

            for source_row in self.click_order_sequence:
                row_items = []

                for col in range(self.table.columnCount()):
                    item = self.table.item(source_row, col)
                    row_items.append(item.clone() if item else QTableWidgetItem(""))

                ordered_rows.append(row_items)

            self.table.setRowCount(0)

            for row_items in ordered_rows:
                row = self.table.rowCount()
                self.table.insertRow(row)

                for col, item in enumerate(row_items):
                    self.table.setItem(row, col, item)

            self.click_order_mode = False
            self.click_order_sequence = []
            self.btn_click_order.setText("Click Order Mode")

        finally:
            self.updating_table = False

        self.auto_fill_from_row(0)
        self._log("[Tab7] Click ordering applied successfully.")

    def choose_color(self, row: int, col: int):
        if self.click_order_mode:
            return

        if col not in (3, 4):
            return

        modes = ["Dark", "Light", "Custom"]
        current_mode = self.table.item(row, 3).text().strip()

        mode, ok = QInputDialog.getItem(
            self,
            "Choose Font Color Mode",
            "Select color mode:",
            modes,
            modes.index(current_mode) if current_mode in modes else 0,
            False,
        )

        if not ok:
            return

        self.updating_table = True

        try:
            if mode == "Dark":
                color_hex = rgb_to_hex(FONT_COLOR)

            elif mode == "Light":
                color_hex = rgb_to_hex(FONT_COLOR_LIGHT)

            else:
                current_color = QColor(self.table.item(row, 4).text().strip())

                color = QColorDialog.getColor(
                    current_color,
                    self,
                    "Choose Custom Page Number Font Color",
                )

                if not color.isValid():
                    return

                color_hex = color.name().upper()

            self.table.item(row, 3).setText(mode)

            custom_item = self.table.item(row, 4)
            custom_item.setText(color_hex)
            self.apply_color_preview(custom_item, color_hex)

        finally:
            self.updating_table = False

    def on_item_changed(self, item):
        if self.updating_table:
            return

        if item.column() == 2:
            self.auto_fill_from_row(item.row())

    def auto_fill_from_row(self, start_row: int = 0):
        if self.table.rowCount() == 0:
            return

        try:
            self.updating_table = True
            page_no = int(self.table.item(start_row, 2).text().strip())

            if page_no < 1:
                raise ValueError("Start page number must be greater than 0.")

            for row in range(start_row, self.table.rowCount()):
                self.table.item(row, 2).setText(str(page_no))
                slides = int(self.table.item(row, 1).text().strip())
                page_no += slides

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

        finally:
            self.updating_table = False

    def _move_up(self):
        if self.click_order_mode:
            self._toggle_click_order_mode()

        row = self.table.currentRow()

        if row <= 0:
            return

        self.table.move_row(row, row - 1)

    def _move_down(self):
        if self.click_order_mode:
            self._toggle_click_order_mode()

        row = self.table.currentRow()

        if row < 0 or row >= self.table.rowCount() - 1:
            return

        self.table.move_row(row, row + 2)

    def get_rows(self) -> list[dict]:
        input_folder = Path(self.input_folder_edit.text().strip())
        rows = []

        for row in range(self.table.rowCount()):
            file_item = self.table.item(row, 0)
            filename = str(file_item.data(Qt.ItemDataRole.UserRole) or file_item.text()).strip()

            slides_text = self.table.item(row, 1).text().strip()
            start_text = self.table.item(row, 2).text().strip()
            color_mode = self.table.item(row, 3).text().strip()
            custom_color = self.table.item(row, 4).text().strip()

            if not slides_text.isdigit():
                raise ValueError(f"Invalid slide count for file: {filename}")

            if not start_text.isdigit() or int(start_text) < 1:
                raise ValueError(f"Invalid start page number for file: {filename}")

            if color_mode not in ["Dark", "Light", "Custom"]:
                raise ValueError(f"Invalid color mode for file: {filename}")

            if color_mode == "Custom" and not QColor(custom_color).isValid():
                raise ValueError(f"Invalid custom color for file: {filename}")

            ppt_path = input_folder / filename

            if not ppt_path.exists():
                raise ValueError(f"File not found: {ppt_path}")

            rows.append(
                {
                    "path": ppt_path,
                    "slides": int(slides_text),
                    "start_page": int(start_text),
                    "color_mode": color_mode,
                    "custom_color": custom_color,
                }
            )

        return rows

    def _run_update(self):
        input_folder = Path(self.input_folder_edit.text().strip())

        if not input_folder.exists() or not input_folder.is_dir():
            QMessageBox.warning(
                self,
                "Invalid Input Folder",
                "Please select a valid input folder.",
            )
            return

        output_folder_text = self.output_folder_edit.text().strip()

        if not output_folder_text:
            QMessageBox.warning(
                self,
                "Invalid Output Folder",
                "Please select a valid output folder.",
            )
            return

        output_folder = Path(output_folder_text)
        output_folder.mkdir(parents=True, exist_ok=True)

        try:
            rows = self.get_rows()

            if not rows:
                raise ValueError("No PPTX files loaded.")

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))
            return

        self._log("=" * 70)
        self._log(f"[Tab7] Input folder:  {input_folder}")
        self._log(f"[Tab7] Output folder: {output_folder}")
        self._log(f"[Tab7] Rows loaded:   {len(rows)}")
        self._log(f"[Tab7] Index table object name: {INDEX_TABLE_OBJECT_NAME}")
        self._log("[Tab7] Index page number will update in Column 2.")
        self._log("[Tab7] Starting reference/page-number update...")

        self.btn_run.setEnabled(False)
        self._set_progress_busy("Updating references and page numbers...")

        worker = FunctionWorker(
            update_references_and_page_numbers,
            rows,
            output_folder,
        )

        worker.signals.started.connect(lambda: self._log("[Tab7] Worker started"))
        worker.signals.progress.connect(self._on_progress)
        worker.signals.finished.connect(self._on_finished)
        worker.signals.error.connect(self._on_error)

        self.thread_pool.start(worker)

    def _on_finished(self, result):
        self.btn_run.setEnabled(True)
        output_folder = Path(result)

        if self.state is not None and hasattr(self.state, "tab7_output_folder"):
            self.state.tab7_output_folder = output_folder

        self._set_progress_value(100, "Completed")
        self._log(f"[Tab7] Updated PPTX folder: {output_folder}")

        QMessageBox.information(
            self,
            "Update Complete",
            f"References + page numbers updated.\n\nOutput folder:\n{output_folder}",
        )

    def _on_error(self, error_text: str):
        self.btn_run.setEnabled(True)
        self._set_progress_idle()
        self.progress_status_label.setText("Failed")

        self._log("[Tab7] ERROR")
        self._log(error_text)

        QMessageBox.critical(self, "Update Failed", error_text)


# Friendly aliases if loaded by dynamic tab loader.
UpdateRefsStandaloneWidget = UpdateRefsWidget
Tab5Widget = UpdateRefsWidget
Tab7Widget = UpdateRefsWidget


def main():
    app = QApplication(sys.argv)

    window = UpdateRefsWidget()
    window.resize(1350, 760)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()