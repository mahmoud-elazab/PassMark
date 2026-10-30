"""
Read an optional reference Excel that has the same FT Data Summary layout as our output.
Extract:
- Unloaded Current (mA) column (single column)
- Loaded Current (A) columns per operating mode (one per mode section)

Matching rule:
- Values are matched by row order starting from the first data row (below 'Sample #').
- Extra reference rows are ignored.
- Missing reference rows yield blanks in output.
"""

from __future__ import annotations
from typing import Dict, List, Optional, Tuple

import openpyxl


def _s(v) -> str:
    return "" if v is None else str(v).strip()


def _norm(v) -> str:
    return _s(v).lower().replace("\n", " ").replace("\r", " ")


def _is_probable_sample_id(v: str) -> bool:
    """
    Your sample IDs look numeric (e.g., 339943).
    We'll treat a row as data if Sample cell is all digits.
    """
    v = v.strip()
    return v.isdigit()


def _find_sample_header_cell(ws) -> Tuple[int, int]:
    """
    Finds the cell containing 'Sample #' anywhere in the sheet.
    Returns (row, col).
    """
    for row in ws.iter_rows():
        for cell in row:
            if _norm(cell.value) == "sample #":
                return cell.row, cell.column
    raise ValueError("Could not find 'Sample #' header in reference Excel.")


def _find_col_by_contains(ws, rows_to_check: List[int], needle: str) -> Optional[int]:
    needle = needle.lower()
    for r in rows_to_check:
        if r < 1:
            continue
        for c in range(1, ws.max_column + 1):
            if needle in _norm(ws.cell(r, c).value):
                return c
    return None


def _find_unloaded_col(ws, header_row: int) -> int:
    # In our layout, Unloaded header is merged across row 4-5, so it may live on header_row-1.
    col = _find_col_by_contains(ws, [header_row, header_row - 1], "unloaded current")
    if col:
        return col

    # Fallback: scan whole sheet
    for row in ws.iter_rows():
        for cell in row:
            if "unloaded current" in _norm(cell.value):
                return cell.column

    raise ValueError("Could not find 'Unloaded Current (mA)' header in reference Excel.")


def _find_loaded_cols_by_mode(ws, header_row: int, mode_label_row: int) -> Dict[int, int]:
    """
    Find columns containing 'Loaded Current' and map them to internal mode numbers.

    This template places the mode label (e.g., '30W/30W') only in the first column
    of each block:
        [Port 1] [Port 2] [Loaded Current]
    Therefore, for a loaded-current column, we scan left to find the mode label.
    """
    label_to_mode = {
        "45w/15w": 2,
        "15w/45w": 3,
        "15w/15w": 4,
        "30w/30w": 5,
        "0w/60w": 6,
        "60w/0w": 7,
    }

    def is_loaded_header(val) -> bool:
        s = _norm(val)
        return ("loaded" in s) and ("current" in s)  # matches "Loaded Current (A)"

    # 1) Collect all "Loaded Current" columns by looking at the header row
    loaded_cols = []
    for c in range(1, ws.max_column + 1):
        if is_loaded_header(ws.cell(header_row, c).value):
            loaded_cols.append(c)

    # 2) For each loaded-current column, find the mode label by scanning left
    out: Dict[int, int] = {}
    for c in loaded_cols:
        mode_label = _norm(ws.cell(mode_label_row, c).value)

        # Mode label is often at the start of a 3-col block; scan left up to 3 cols
        if not mode_label:
            for back in range(1, 4):
                cc = c - back
                if cc < 1:
                    break
                candidate = _norm(ws.cell(mode_label_row, cc).value)
                if candidate in label_to_mode:
                    mode_label = candidate
                    break

        if mode_label in label_to_mode:
            out[label_to_mode[mode_label]] = c

    return out


def _find_data_start_row(ws, header_row: int, sample_col: int) -> int:
    """
    Data doesn't always start immediately at header_row+1 (sometimes there are blanks).
    Find the first row below header where sample looks like numeric ID.
    """
    r = header_row + 1
    while r <= ws.max_row:
        v = _s(ws.cell(r, sample_col).value)
        if _is_probable_sample_id(v):
            return r
        # Stop early if we hit Criteria immediately (unlikely but safe)
        if "criteria" in _norm(v):
            break
        r += 1
    # fallback to header_row+1 if nothing found
    return header_row + 1


def _read_down_until_criteria(ws, start_row: int, sample_col: int, value_col: int) -> List[Optional[float]]:
    vals: List[Optional[float]] = []
    r = start_row
    while r <= ws.max_row:
        sample_raw = _s(ws.cell(r, sample_col).value)
        sample_n = _norm(sample_raw)

        if sample_n == "" or "criteria" in sample_n:
            break

        v = ws.cell(r, value_col).value
        if v is None or _s(v) == "":
            vals.append(None)
        else:
            try:
                vals.append(float(v))
            except Exception:
                vals.append(None)
        r += 1

    return vals


def read_reference_currents(filepath: str, sheet_name: Optional[str] = None) -> Dict:
    """
    Reads the optional reference Excel and extracts:
      - unloaded current (single column)
      - loaded current (multiple columns, mapped to operating modes by the mode label row)

    Also prints debug output for loaded-current detection and values.

    Returns:
      {
        'unloaded_mA': [...],
        'loaded_A_by_mode': {2:[...],3:[...],...},
        'source_rows': int,
        'debug': {...}
      }
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # If a sheet name is provided, use it; else prefer "FT Data Summary" if present.
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb["FT Data Summary"] if "FT Data Summary" in wb.sheetnames else wb.worksheets[0]

    header_row, sample_col = _find_sample_header_cell(ws)
    mode_label_row = header_row - 1

    data_start_row = _find_data_start_row(ws, header_row, sample_col)

    unloaded_col = _find_unloaded_col(ws, header_row)
    loaded_cols_by_mode = _find_loaded_cols_by_mode(ws, header_row, mode_label_row)

    unloaded_mA = _read_down_until_criteria(ws, data_start_row, sample_col, unloaded_col)

    loaded_A_by_mode: Dict[int, List[Optional[float]]] = {}
    for mode, col in loaded_cols_by_mode.items():
        loaded_A_by_mode[mode] = _read_down_until_criteria(ws, data_start_row, sample_col, col)

    return {
        "unloaded_mA": unloaded_mA,
        "loaded_A_by_mode": loaded_A_by_mode,
        "source_rows": len(unloaded_mA),
        "debug": {
            "sheet": ws.title,
            "header_row": header_row,
            "mode_label_row": mode_label_row,
            "data_start_row": data_start_row,
            "sample_col": sample_col,
            "unloaded_col": unloaded_col,
            "loaded_cols_by_mode": loaded_cols_by_mode,
        }
    }