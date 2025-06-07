# Utility helper functions for Excel MCP
from typing import Dict, List, Any
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter


def _col_to_index(col: str) -> int:
    """Convert Excel column letters (e.g. 'A', 'BC') to a 1-based index."""
    idx = 0
    for ch in col.upper():
        if not ch.isalpha():
            continue
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx


def address_within_ranges(target: str, ranges: List[str]) -> bool:
    """Check if an address is contained within any sheet range.

    Parameters
    ----------
    target:
        Address like ``Sheet1!A1``.
    ranges:
        List of range addresses such as ``Sheet1!A1:C10``.

    Returns
    -------
    bool
        ``True`` if ``target`` falls within one of ``ranges``.
    """
    if '!' not in target or not ranges:
        return False

    tgt_sheet, tgt_cell = target.split('!')
    tgt_col = ''.join(filter(str.isalpha, tgt_cell))
    tgt_row = int(''.join(filter(str.isdigit, tgt_cell)))
    tgt_col_idx = _col_to_index(tgt_col)

    for entry in ranges:
        if '!' not in entry:
            continue
        range_sheet, range_part = entry.split('!')
        if range_sheet != tgt_sheet or ':' not in range_part:
            continue
        start, end = range_part.split(':')
        start_col = ''.join(filter(str.isalpha, start))
        start_row = int(''.join(filter(str.isdigit, start)))
        end_col = ''.join(filter(str.isalpha, end))
        end_row = int(''.join(filter(str.isdigit, end)))
        if start_row <= tgt_row <= end_row:
            start_idx = _col_to_index(start_col)
            end_idx = _col_to_index(end_col)
            if start_idx <= tgt_col_idx <= end_idx:
                return True
    return False


def collect_column_outputs(cells: Dict[str, Dict[str, Any]], anchor: str, text_limit: int = 3) -> Dict[str, Any]:
    """Gather output values from cells in the same column as ``anchor``.

    Scans upward until ``text_limit`` consecutive non-formula and non-numeric
    cells are encountered, then scans downward until the next such cell or up
    to 10 rows. Only addresses present in ``cells`` are returned.

    Parameters
    ----------
    cells:
        Mapping of addresses to dictionaries with at least an ``output`` key and
        optionally a ``formula`` key.
    anchor:
        Address like ``A10`` that serves as the starting point.
    text_limit:
        Number of consecutive text cells allowed when scanning upward.

    Returns
    -------
    Dict[str, Any]
        Addresses mapped to their output values.
    """

    row, col_idx = coordinate_to_tuple(anchor)
    column = get_column_letter(col_idx)

    valid = {k: v for k, v in cells.items() if v.get("output") is not None}

    result: Dict[str, Any] = {}

    consecutive_text = 0
    current = row - 1
    inspected = 0
    while current >= 1 and inspected < 100:
        addr = f"{column}{current}"
        if addr in valid:
            info = valid[addr]
            result[addr] = info["output"]
            inspected += 1
            if "formula" not in info:
                try:
                    float(info["output"])
                    is_number = True
                except (ValueError, TypeError):
                    is_number = False
                if not is_number:
                    consecutive_text += 1
                    if consecutive_text >= text_limit:
                        break
        current -= 1

    if anchor in cells:
        result[anchor] = cells[anchor].get("output")

    max_row = max((coordinate_to_tuple(a)[0] for a in valid), default=row)
    current = row + 1
    added = 0
    while current <= max_row and added < 10:
        addr = f"{column}{current}"
        if addr in valid:
            info = valid[addr]
            result[addr] = info["output"]
            added += 1
            if "formula" not in info:
                try:
                    float(info["output"])
                    is_number = True
                except (ValueError, TypeError):
                    is_number = False
                if not is_number:
                    break
        current += 1

    return result
