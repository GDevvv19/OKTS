"""
Parse the OK-TO-SHIP workbook to learn which units ship and which are held back.

Two supported layouts (auto-detected):

A) Per-IMEI "Shipped Status" column  (preferred when present)
   The team fills in our template and edits a "Shipped Status" column (TRUE/FALSE
   or Yes/No) next to each IMEI. We read IMEI -> Shipped Status directly.

B) Two tables in the "Log file verification" tab  (original layout)
   * "Count of <field>"   — a pivot table giving the full lot of units (the
     shippable set). The field name varies by how the pivot was built (seen
     as both "Count of STATION" and "Count of IMEI"); only the "Count of"
     prefix is matched, not the field name.
   * "Balance SN status"  — the units NOT being shipped (held back;
     also "Balance status" / "Balance SN"). A unit is:
         in Balance table       -> FALSE (not shipped)
         else in Count of table -> TRUE  (shipped)

Anything not found in the OK-TO-SHIP -> None (unknown). Header wording/positions
vary, so everything is located by fuzzy text matching, not fixed cells.
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

_ID_RE = re.compile(r"^[0-9A-Za-z]{10,32}$")   # IMEI (digits) or DevEUI (hex)
_SKIP_WORDS = {"qty", "remarks", "fct result", "status", "row labels",
               "column labels", "grand total", "count of station",
               "balance status", "balance sn status", "balance sn", "balance",
               "pass", "fail", "sn", "imei", "serial", "serial number"}


def _norm_id(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _looks_like_id(v) -> bool:
    s = _norm_id(v)
    return bool(_ID_RE.match(s)) and any(ch.isdigit() for ch in s) \
        and s.lower() not in _SKIP_WORDS


def _find_sheet(wb, wanted="Log file verification"):
    norm = wanted.strip().lower()
    for name in wb.sheetnames:
        if name.strip().lower() == norm:
            return wb[name]
    for name in wb.sheetnames:
        if norm in name.strip().lower():
            return wb[name]
    raise KeyError(f"Sheet '{wanted}' not found. Available: {wb.sheetnames}")


_COUNT_OF_RE = re.compile(r"\bcount of\s+\S+", re.IGNORECASE)


def _find_count_of_header(ws, max_row=12, max_col=60):
    """Find the pivot table's 'Count of <field>' header — this is Excel's
    auto-generated label for whichever field the pivot happened to be built
    from (seen as both "Count of STATION" and "Count of IMEI" across
    different OK-TO-SHIP template versions), but it always marks the SAME
    thing: the full shippable lot. Matching the fixed trailing word broke the
    moment a template used a different field name — matching the "Count of"
    pattern itself works regardless of which field name Excel chose."""
    for r in range(1, min(ws.max_row, max_row) + 1):
        for c in range(1, min(ws.max_column, max_col) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and _COUNT_OF_RE.search(v):
                return r, c
    return None


def _find_header(ws, *needles, max_row=12, max_col=60):
    """Return (row, col) of the first cell whose text CONTAINS any needle
    (case-insensitive). Needles are tried in order."""
    for needle in needles:
        t = needle.strip().lower()
        for r in range(1, min(ws.max_row, max_row) + 1):
            for c in range(1, min(ws.max_column, max_col) + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and t in v.strip().lower():
                    return r, c
    return None


def _collect_ids_below(ws, header_row, header_col, spread=3, max_gap=3):
    """Collect id-like values in the columns at/after a table header, scanning
    downward. Tolerates a couple of sub-header rows and small blank gaps."""
    ids: set[str] = set()
    for col in range(header_col, header_col + spread + 1):
        gap = 0
        found_any = False
        for r in range(header_row + 1, ws.max_row + 1):
            v = ws.cell(r, col).value
            if v in (None, ""):
                gap += 1
                if found_any and gap > max_gap:
                    break
                continue
            gap = 0
            if _looks_like_id(v):
                ids.add(_norm_id(v))
                found_any = True
            elif str(v).strip().lower() == "grand total":
                break
        if found_any:
            break   # the first column that actually holds ids is the right one
    return ids


_SHIP_TRUE = {"true", "yes", "y", "1", "ship", "shipped", "ok", "okts",
              "ok to ship", "oktoship"}
_SHIP_FALSE = {"false", "no", "n", "0", "hold", "held", "balance", "bal",
               "not shipped", "reject", "rejected"}


def _interpret_ship(v):
    """Map a Shipped-Status cell value to True / False / None."""
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    s = str(v).strip().lower()
    if s in _SHIP_TRUE:
        return True
    if s in _SHIP_FALSE:
        return False
    if "balance" in s or "hold" in s or "not ship" in s:
        return False
    if "ship" in s or "okts" in s:
        return True
    return None


def _find_status_table(wb):
    """Layout A: find a header row with an IMEI column and a 'Shipped Status'
    column (on any sheet) and return {imei: True/False}. {} if not present."""
    for ws in wb.worksheets:
        max_r = min(ws.max_row, 20)
        max_c = min(ws.max_column, 200)
        for hr in range(1, max_r + 1):
            status_col = imei_col = None
            for c in range(1, max_c + 1):
                h = str(ws.cell(hr, c).value or "").strip().lower()
                if not h:
                    continue
                if status_col is None and ("shipped status" in h or h == "shipped"):
                    status_col = c
                if imei_col is None and ("imei" in h or "deveui" in h):
                    imei_col = c
            if status_col and imei_col:
                m = {}
                for r in range(hr + 1, ws.max_row + 1):
                    imei = _norm_id(ws.cell(r, imei_col).value)
                    if not imei or not any(ch.isdigit() for ch in imei):
                        continue
                    b = _interpret_ship(ws.cell(r, status_col).value)
                    if b is not None:
                        m[imei] = b
                if m:
                    return m
    return {}


class OkToShip:
    def __init__(self, all_imeis: set[str], balance_imeis: set[str], status_map=None):
        self.all_imeis = all_imeis          # Count of STATION (full lot)
        self.balance_imeis = balance_imeis  # units NOT shipping (held back)
        self.status_map = status_map or {}  # explicit per-IMEI Shipped Status

    def shipped_status(self, imei: str):
        """True (ship), False (held back), or None (unknown to OK-TO-SHIP)."""
        imei = _norm_id(imei)
        if self.status_map:                 # layout A takes precedence
            return self.status_map.get(imei)
        if imei in self.balance_imeis:
            return False
        if imei in self.all_imeis:
            return True
        return None


def parse_ok_to_ship(path: Path) -> OkToShip:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        # Layout A — explicit per-IMEI "Shipped Status" column
        status_map = _find_status_table(wb)
        if status_map:
            ship = {i for i, b in status_map.items() if b}
            hold = {i for i, b in status_map.items() if not b}
            return OkToShip(ship, hold, status_map=status_map)

        # Layout B — "Count of STATION" + "Balance …" tables
        try:
            ws = _find_sheet(wb)
        except KeyError:
            return OkToShip(set(), set())   # unknown layout -> mark nothing

        all_imeis: set[str] = set()
        cos = _find_count_of_header(ws)
        if cos:
            all_imeis = _collect_ids_below(ws, cos[0], cos[1])

        balance: set[str] = set()
        bal = _find_header(ws, "balance sn status", "balance sn",
                           "balance status", "balance")
        if bal:
            balance = _collect_ids_below(ws, bal[0], bal[1])

        return OkToShip(all_imeis, balance)
    finally:
        wb.close()
