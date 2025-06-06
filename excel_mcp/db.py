import duckdb
import time
from typing import Dict, Optional, List, Tuple

_db_conn: Optional[duckdb.DuckDBPyConnection] = None


def init_db(path: str = "excel_mcp.db") -> None:
    """Initialize DuckDB connection and create tables if needed."""
    global _db_conn
    _db_conn = duckdb.connect(path)
    _db_conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cell_labels(
            label TEXT,
            sheet_name TEXT,
            cell_address TEXT,
            last_updated TIMESTAMP,
            PRIMARY KEY(label, sheet_name)
        )
        """
    )


def store_label_map(sheet_name: str, label_map: Dict[str, str]) -> None:
    """Insert or update label mappings in the database."""
    if _db_conn is None:
        return
    ts = time.time()
    for label, addr in label_map.items():
        _db_conn.execute(
            "DELETE FROM cell_labels WHERE label = ? AND sheet_name = ?",
            (label, sheet_name),
        )
        _db_conn.execute(
            "INSERT INTO cell_labels VALUES (?, ?, ?, to_timestamp(?))",
            (label, sheet_name, addr, ts),
        )


def query_label(label: str) -> List[Tuple[str, str]]:
    """Return list of (sheet_name, cell_address) for a label."""
    if _db_conn is None:
        return []
    rows = _db_conn.execute(
        "SELECT sheet_name, cell_address FROM cell_labels WHERE label = ?",
        (label,),
    ).fetchall()
    return [(r[0], r[1]) for r in rows]
