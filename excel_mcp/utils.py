# Utility helper functions for Excel MCP
from typing import List


def _col_to_index(col: str) -> int:
    """Convert Excel column letters (e.g. 'A', 'BC') to a 1-based index."""
    idx = 0
    for ch in col.upper():
        if not ch.isalpha():
            continue
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx


def _index_to_col(idx: int) -> str:
    """Convert a 1-based column index to Excel column letters."""
    col = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        col = chr(rem + ord("A")) + col
    return col


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


def gather_row_context(rows: dict, start: str, text_limit: int = 3) -> dict:
    """Collect outputs for cells near ``start`` in the same row.

    Parameters
    ----------
    rows:
        Mapping of cell addresses to dictionaries containing at least an
        ``output`` entry and optionally a ``formula`` entry.
    start:
        Address of the reference cell (e.g. ``"B2"``).
    text_limit:
        Number of consecutive non-formula, non-numeric cells to allow when
        scanning left from ``start``.

    Returns
    -------
    dict
        Mapping of addresses to the stored ``output`` values for the scanned
        cells.
    """

    def _is_number(val) -> bool:
        try:
            float(val)
            return True
        except (ValueError, TypeError):
            return False

    start = start.strip("$")
    col_part = "".join(filter(str.isalpha, start))
    row_part = int("".join(filter(str.isdigit, start)))
    start_idx = _col_to_index(col_part)

    with_output = {k: v for k, v in rows.items() if v.get("output") is not None}

    results = {}

    # Scan left from the start cell
    consecutive_text = 0
    idx = start_idx - 1
    checked = 0
    while idx >= 1 and checked < 100:
        addr = f"{_index_to_col(idx)}{row_part}"
        if addr in with_output:
            info = with_output[addr]
            results[addr] = info["output"]
            checked += 1
            if "formula" not in info and not _is_number(info["output"]):
                consecutive_text += 1
                if consecutive_text >= text_limit:
                    break
        idx -= 1

    # Include the start cell if present
    start_cell = f"{col_part}{row_part}"
    if start_cell in rows:
        results[start_cell] = rows[start_cell].get("output")

    # Scan to the right of the start cell
    idx = start_idx + 1
    if rows:
        max_idx = max(_col_to_index("".join(filter(str.isalpha, k))) for k in rows.keys())
    else:
        max_idx = start_idx
    added = 0
    while idx <= max_idx and added < 10:
        addr = f"{_index_to_col(idx)}{row_part}"
        if addr in with_output:
            info = with_output[addr]
            results[addr] = info["output"]
            added += 1
            if "formula" not in info and not _is_number(info["output"]):
                return results
        idx += 1

    return results
