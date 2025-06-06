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
