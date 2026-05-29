"""
Isolates classification logic and boundary detection.
Implements custom formula hypervisor enforcing intra-row recalculations,
headless column pruning, trailing summary truncation, strict AMT calculations,
and contiguous serial re-indexing.
"""
import logging
import re
from typing import Dict, List, Set, Tuple, Any
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.utils import column_index_from_string, get_column_letter

from config import PANEL_ROW, HEADER_ROW, PANEL_START_COL, END_MARKER

def _is_empty(val: Any) -> bool:
    """Evaluates if a cell value constitutes an empty or null equivalent state."""
    if val is None:
        return True
    return str(val).strip().upper() in ("", "0", "0.0", "NONE", "NULL", "-")

def _is_numeric(val: Any) -> bool:
    """Evaluates if a cell value is strictly numerical."""
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return True
    try:
        float(val)
        return True
    except ValueError:
        return False

def _sanitize_sheet_title(title: str) -> str:
    """Strips Excel-forbidden characters and enforces the 31-character limit."""
    safe_title = re.sub(r'[\\/*?:\[\]]', '_', title)
    return safe_title[:31]

def _truncate_trailing_summary(ws: Worksheet, desc_col: int, cat_col: int) -> None:
    """
    Severs trailing summary tables to eliminate circular references.
    Halts and truncates the matrix when two consecutive rows lack both DESCRIPTION and CAT NO.
    """
    consecutive_blanks = 0
    cutoff_row = -1

    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        desc_val = ws.cell(row=r, column=desc_col).value
        cat_val = ws.cell(row=r, column=cat_col).value

        if _is_empty(desc_val) and _is_empty(cat_val):
            consecutive_blanks += 1
            if consecutive_blanks == 2:
                cutoff_row = r - 1
                break
        else:
            consecutive_blanks = 0

    if cutoff_row != -1:
        rows_to_delete = ws.max_row - cutoff_row + 1
        if rows_to_delete > 0:
            ws.delete_rows(cutoff_row, rows_to_delete)

def _get_shifted_merges(ws: Worksheet, cols_to_delete: List[int]) -> List[CellRange]:
    """Calculates the exact new position for merged cells ignoring openpyxl bugs."""
    new_ranges = []
    for mr in list(ws.merged_cells.ranges):
        shift_min = sum(1 for c in cols_to_delete if c < mr.min_col)
        shift_max = sum(1 for c in cols_to_delete if c < mr.max_col)

        new_min_col = mr.min_col - shift_min
        new_max_col = mr.max_col - shift_max

        if new_min_col <= new_max_col:
            new_mr = CellRange(min_col=new_min_col, max_col=new_max_col,
                               min_row=mr.min_row, max_row=mr.max_row)
            new_ranges.append(new_mr)
    return new_ranges

def _shift_formulas_and_unhide(ws: Worksheet, cols_to_delete: List[int]) -> None:
    """
    Safely un-hides formulas, applies absolute column translation mapping,
    and enforces strict intra-row recalculations by anchoring row numbers to the current row.
    """
    shift_map = {}
    for col_idx in range(1, 2000):
        shift = sum(1 for c in cols_to_delete if c < col_idx)
        shift_map[col_idx] = col_idx - shift

    pattern = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')

    for row in ws.iter_rows():
        current_row = row[0].row

        def replacer(match) -> str:
            abs_col = match.group(1)
            col_letters = match.group(2)
            abs_row = match.group(3)

            try:
                col_idx = column_index_from_string(col_letters)
            except ValueError:
                return match.group(0)

            if col_idx in cols_to_delete:
                return "#REF!"

            new_col_idx = shift_map.get(col_idx, col_idx)
            new_col_letters = get_column_letter(new_col_idx)

            return f"{abs_col}{new_col_letters}{abs_row}{current_row}"

        for cell in row:
            if isinstance(cell.value, str):
                if cell.value.startswith('FORMULA_HIDE='):
                    orig_formula = cell.value.replace('FORMULA_HIDE', '', 1)
                    cell.value = pattern.sub(replacer, orig_formula)
                elif cell.value.startswith("'="):
                    orig_formula = cell.value.replace("'", "", 1)
                    cell.value = pattern.sub(replacer, orig_formula)

def _fix_amt_formulas(ws: Worksheet) -> None:
    """
    Dynamically overwrites 'AMT' formulas to mandate the multiplication of
    the immediate left column by Column F (QTY).
    """
    amt_cols = []
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(row=HEADER_ROW, column=c).value or "").strip().upper() == "AMT":
            amt_cols.append(c)

    if not amt_cols:
        return

    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        qty_val = ws.cell(row=r, column=6).value

        if not _is_empty(qty_val):
            for c in amt_cols:
                left_col_letter = get_column_letter(c - 1)
                ws.cell(row=r, column=c).value = f"={left_col_letter}{r}*F{r}"

def _clear_headless_columns(ws: Worksheet) -> None:
    """
    Scans the first 5 rows of each column. If no label exists,
    the entire column is systematically cleared of data.
    """
    for c in range(1, ws.max_column + 1):
        has_header = False
        for r in range(1, 6):
            if not _is_empty(ws.cell(row=r, column=c).value):
                has_header = True
                break

        if not has_header:
            for r in range(1, ws.max_row + 1):
                ws.cell(row=r, column=c).value = None

def _reindex_serial_numbers(ws: Worksheet, sno_col: int) -> None:
    """
    Enforces contiguous sequential numbering for active data rows post-pruning.
    Overrides existing numerical entries to ensure logical output structure.
    """
    current_sno = 1
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        sno_val = ws.cell(row=r, column=sno_col).value
        if _is_numeric(sno_val):
            ws.cell(row=r, column=sno_col).value = current_sno
            current_sno += 1

def detect_columns(ws: Worksheet, logger: logging.Logger) -> Tuple[Dict[str, List[int]], List[int], Dict[str, int]]:
    panel_cols: Dict[str, List[int]] = {}
    all_panel_cols: List[int] = []
    header_map: Dict[str, int] = {}

    end_col: int = -1
    for col in range(PANEL_START_COL, ws.max_column + 1):
        val1 = str(ws.cell(row=1, column=col).value or "").strip().upper()
        val2 = str(ws.cell(row=PANEL_ROW, column=col).value or "").strip().upper()
        if END_MARKER in val1 or END_MARKER in val2:
            end_col = col
            break

    if end_col == -1:
        logger.warning(f"Boundary marker '{END_MARKER}' not found. Scanning to max boundary.")
        end_col = ws.max_column + 1

    for col in range(PANEL_START_COL, end_col):
        raw_val = ws.cell(row=PANEL_ROW, column=col).value
        val = str(raw_val or "").strip()
        if val:
            all_panel_cols.append(col)
            prefix = val.split("-")[0].strip()
            if prefix not in panel_cols:
                panel_cols[prefix] = []
            panel_cols[prefix].append(col)

    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=HEADER_ROW, column=col).value
        if val:
            header_map[str(val).strip().upper()] = col

    return panel_cols, all_panel_cols, header_map

def classify_panels(wb: Workbook, ws: Worksheet, panel_cols: Dict[str, List[int]],
                   all_panel_cols: List[int], header_map: Dict[str, int], logger: logging.Logger) -> int:
    sheets_created: int = 0
    sno_col: int = header_map.get('SNO', 1)
    desc_col: int = header_map.get('DESCRIPTION', 2)
    cat_col: int = header_map.get('CAT NO.', 6)

    for prefix, prefix_panel_cols in panel_cols.items():
        logger.info(f"Generating isolated structure for series: {prefix}")

        ws_new: Worksheet = wb.copy_worksheet(ws)
        ws_new.title = _sanitize_sheet_title(prefix)

        _truncate_trailing_summary(ws_new, desc_col, cat_col)

        data_rows_to_delete: Set[int] = set()

        for r in range(HEADER_ROW + 1, ws_new.max_row + 1):
            sno_val = ws_new.cell(row=r, column=sno_col).value
            cat_val = ws_new.cell(row=r, column=cat_col).value

            is_data_row = _is_numeric(sno_val) or not _is_empty(cat_val)
            if is_data_row:
                has_prefix_qty = any(
                    not _is_empty(ws_new.cell(row=r, column=c).value) and ws_new.cell(row=r, column=c).value != 0
                    for c in prefix_panel_cols
                )
                if not has_prefix_qty:
                    data_rows_to_delete.add(r)

        rows_to_delete: List[int] = sorted(list(data_rows_to_delete), reverse=True)
        for r in rows_to_delete:
            ws_new.delete_rows(r)

        cols_to_delete: List[int] = sorted([c for c in all_panel_cols if c not in prefix_panel_cols], reverse=True)

        for row in ws_new.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = "FORMULA_HIDE" + cell.value

        shifted_merges = _get_shifted_merges(ws_new, cols_to_delete)
        ws_new.merged_cells.ranges = []

        for c in cols_to_delete:
            ws_new.delete_cols(c)

        _shift_formulas_and_unhide(ws_new, cols_to_delete)

        _fix_amt_formulas(ws_new)

        for mr in shifted_merges:
            ws_new.merged_cells.add(mr)

        _clear_headless_columns(ws_new)

        # Shift target sno column index if preceding columns were deleted
        actual_sno_col = sno_col - sum(1 for c in cols_to_delete if c < sno_col)
        _reindex_serial_numbers(ws_new, actual_sno_col)

        sheets_created += 1

    return sheets_created