"""
Fill a COPY of the sample template with computed ReportData, preserving the
template's formatting, formulas, and charts.

Strategy: the workbook is a byte-copy of the sample, so every style / formula /
chart already exists. We only (a) overwrite header/summary cells, and (b) write
the variable-length data tables — extending or trimming rows and cloning the
style of the template's first data row so added rows look identical.
"""

from __future__ import annotations

import re
from copy import copy
from typing import Optional

from openpyxl.styles import PatternFill, Font, Border, Alignment

from analysis import ReportData, Device
from ok_to_ship import OkToShip

_BLANK_FILL = PatternFill(fill_type=None)   # shared "no fill" for clearing cells
_BLANK_BORDER = Border()                    # shared "no border"
_DEFAULT_FONT = Font()                      # shared default font
_DEFAULT_ALIGN = Alignment()                # shared default alignment
_SKIP = object()                            # sentinel: "leave this column as-is"

# Column letter -> index helpers are avoided; we use numeric columns directly.
A, B, C, D, E, F, G, H, I, J = 1, 2, 3, 4, 5, 6, 7, 8, 9, 10

RAW_ATTEMPTS_BLOCKS = 5     # Test 1..5 in (Raw) Attempts
MC_BLOCKS = 6               # Test 1..6 in MC Retest Units
BLOCK_COLS = 6              # Tester, Station, Date, Time, Result, Fail Reason


def _copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def _snap(cell):
    """Snapshot a cell's visual style so it can be re-applied elsewhere."""
    return (copy(cell.font), copy(cell.fill), copy(cell.border),
            copy(cell.alignment), cell.number_format)


def _apply_snap(cell, snap):
    if snap is None:
        return
    cell.font, cell.fill, cell.border, cell.alignment, cell.number_format = snap


def _capture_block_styles(ws, start, block_start=7, scan=1000):
    """From the template's data rows, grab a representative 6-cell Test-block
    style for a PASS attempt, a FAIL attempt, and an empty block — so generated
    rows keep PASS-green / FAIL-red / blank formatting per attempt result."""
    b1, res_col, b2 = block_start, block_start + 4, block_start + BLOCK_COLS
    pass_s = fail_s = empty_s = None
    for r in range(start, start + scan):
        if ws.cell(r, 1).value in (None, ""):
            break
        res = ws.cell(r, res_col).value
        if res == "PASS" and pass_s is None:
            pass_s = [_snap(ws.cell(r, b1 + k)) for k in range(BLOCK_COLS)]
        elif res == "FAIL" and fail_s is None:
            fail_s = [_snap(ws.cell(r, b1 + k)) for k in range(BLOCK_COLS)]
        if empty_s is None and ws.cell(r, b2).value in (None, ""):
            empty_s = [_snap(ws.cell(r, b2 + k)) for k in range(BLOCK_COLS)]
        if pass_s and fail_s and empty_s:
            break
    return pass_s, fail_s, empty_s


def _capture_row(ws, row, max_col):
    """Copy a template row's per-column styles ONCE (returns shared objects)."""
    snap = []
    for c in range(1, max_col + 1):
        cell = ws.cell(row, c)
        if cell.has_style:
            snap.append((copy(cell.font), copy(cell.fill), copy(cell.border),
                         copy(cell.alignment), cell.number_format, copy(cell.protection)))
        else:
            snap.append(None)
    return snap


def _apply_row(ws, row, snap):
    """Assign captured style objects (by reference — no per-cell copy)."""
    for i, s in enumerate(snap):
        if s is None:
            continue
        cell = ws.cell(row, i + 1)
        cell.font, cell.fill, cell.border, cell.alignment, cell.number_format, cell.protection = s


def _clone_row_style(ws, template_row, target_row, max_col):
    if template_row == target_row:
        return
    # cache the captured template-row style on the worksheet (fresh per run) so
    # we copy the style objects once, then reuse them for every data row.
    try:
        cache = ws._rs_cache
    except AttributeError:
        cache = ws._rs_cache = {}
    key = (template_row, max_col)
    snap = cache.get(key)
    if snap is None:
        snap = cache[key] = _capture_row(ws, template_row, max_col)
    _apply_row(ws, target_row, snap)
    ws.row_dimensions[target_row].height = ws.row_dimensions[template_row].height


def _clear_cells(ws, row, max_col):
    """Fully blank a row: clear the value AND remove any leftover template
    formatting (fill/border/font/alignment) so tables end exactly at the data."""
    for c in range(1, max_col + 1):
        if _is_merged_nonanchor(ws, row, c):
            continue
        cell = ws.cell(row, c)
        cell.value = None
        cell.fill = _BLANK_FILL
        cell.border = _BLANK_BORDER
        cell.font = _DEFAULT_FONT
        cell.alignment = _DEFAULT_ALIGN
        cell.number_format = "General"


def _is_merged_nonanchor(ws, row, col):
    """True if (row,col) is inside a merged range but NOT its top-left anchor."""
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            return not (row == mr.min_row and col == mr.min_col)
    return False


def _set(ws, row, col, value):
    # Non-anchor cells of a merged range are read-only; writing them would raise.
    if _is_merged_nonanchor(ws, row, col):
        return
    ws.cell(row, col).value = value


def _formula_at(ws, row, col):
    """Return the formula string at (row,col) if the cell holds one, else None."""
    v = ws.cell(row, col).value
    return v if isinstance(v, str) and v.startswith("=") else None


def _propagate_formula(ws, template_formula, template_row, row, col):
    """Write `template_formula` (captured from template_row) into (row,col),
    translating its relative cell references down to the new row — so we reuse
    the SAMPLE's own formula instead of a hard-coded one."""
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter
    origin = f"{get_column_letter(col)}{template_row}"
    dest = f"{get_column_letter(col)}{row}"
    try:
        ws.cell(row, col).value = Translator(template_formula, origin=origin).translate_formula(dest)
    except Exception:
        ws.cell(row, col).value = template_formula


def _insert_rows_keep_merges(ws, at_row, n):
    """Insert n rows at at_row, correctly moving merged ranges at/below it.
    openpyxl.insert_rows does not touch merged ranges, so we unmerge the affected
    ones first, insert, then re-merge them shifted down by n."""
    if n <= 0:
        return
    from openpyxl.utils import get_column_letter
    affected = [(mr.min_row, mr.max_row, mr.min_col, mr.max_col)
                for mr in list(ws.merged_cells.ranges) if mr.min_row >= at_row]
    for r1, r2, c1, c2 in affected:
        ws.unmerge_cells(f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}")
    ws.insert_rows(at_row, n)
    for r1, r2, c1, c2 in affected:
        ws.merge_cells(f"{get_column_letter(c1)}{r1 + n}:{get_column_letter(c2)}{r2 + n}")


_DEV_ID_HEADER = "Device ID (Dev:)"


import re as _re
_CELL_REF_RE = _re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')


def _shift_formula_cols(formula, at_col, n):
    """Increment every column reference >= at_col by n. openpyxl's
    insert_cols relocates formula CELLS but never rewrites the formula TEXT
    (unlike Excel's own column insert), so a formula that used to read
    B4/C4/D4 would otherwise still say B4/C4/D4 after the columns it
    depends on shifted right."""
    from openpyxl.utils import get_column_letter, column_index_from_string

    def repl(m):
        dollar_col, col_letters, dollar_row, row = m.groups()
        col_idx = column_index_from_string(col_letters)
        if col_idx >= at_col:
            col_idx += n
        return f"{dollar_col}{get_column_letter(col_idx)}{dollar_row}{row}"
    return _CELL_REF_RE.sub(repl, formula)


def _insert_col_keep_merges(ws, at_col, n=1):
    """Insert n blank columns at at_col, correctly widening/shifting merged
    ranges, column widths, and formula references. openpyxl.insert_cols moves
    cell values/styles but leaves merged ranges, column_dimensions, and
    formula text untouched, so we fix all three ourselves: a merged range
    that starts at/after at_col shifts right by n; one that straddles at_col
    just grows wider by n; formulas get their column refs bumped by n."""
    if n <= 0:
        return
    from openpyxl.utils import get_column_letter, column_index_from_string
    affected = [(mr.min_row, mr.max_row, mr.min_col, mr.max_col)
                for mr in list(ws.merged_cells.ranges) if mr.max_col >= at_col]
    for r1, r2, c1, c2 in affected:
        ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    widths = {column_index_from_string(letter): dim.width
              for letter, dim in list(ws.column_dimensions.items()) if dim.width is not None}
    for letter in list(ws.column_dimensions.keys()):
        del ws.column_dimensions[letter]
    ws.insert_cols(at_col, n)
    for idx, w in widths.items():
        new_idx = idx + n if idx >= at_col else idx
        ws.column_dimensions[get_column_letter(new_idx)].width = w
    for r1, r2, c1, c2 in affected:
        if c1 >= at_col:
            c1, c2 = c1 + n, c2 + n
        else:
            c2 += n
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                cell.value = _shift_formula_cols(v, at_col, n)


def _header_row_of(ws, aliases, default, scan=6):
    """Find the row (within the first `scan` rows) whose column-A cell looks
    like one of the given header aliases — the header row position varies a
    little between template versions. Falls back to `default` if not found."""
    for r in range(1, scan + 1):
        v = str(ws.cell(r, 1).value or "").strip().lower()
        if any(a in v for a in aliases):
            return r
    return default


def _has_device_id_col(ws, header_row, col):
    h = str(ws.cell(header_row, col).value or "").strip().lower()
    return "device id" in h or "(dev" in h or "dev:" in h


def _ensure_device_id_col(ws, header_row, imei_col):
    """Make sure the column right after `imei_col` is the 'Device ID (Dev:)'
    column, inserting it if this sample template predates that column.
    Returns that column's index (always imei_col + 1)."""
    dev_col = imei_col + 1
    if _has_device_id_col(ws, header_row, dev_col):
        return dev_col
    _insert_col_keep_merges(ws, dev_col, 1)
    _copy_style(ws.cell(header_row, imei_col), ws.cell(header_row, dev_col))
    ws.cell(header_row, dev_col).value = _DEV_ID_HEADER
    from openpyxl.utils import get_column_letter
    imei_width = ws.column_dimensions[get_column_letter(imei_col)].width
    if imei_width:
        ws.column_dimensions[get_column_letter(dev_col)].width = imei_width
    return dev_col


_SOP_HEADER = "SOP Followed"


def _ensure_sop_column(ws, header_row):
    """Make sure there's an 'SOP Followed' column right before 'Shipped
    Status' in (Raw) Attempts, inserting it if this sample template predates
    it. Returns that column's index, or None if the template has no Shipped
    Status column to anchor next to."""
    ship_col = next((c for c in range(1, ws.max_column + 1)
                      if "shipped status" in str(ws.cell(header_row, c).value or "").lower()
                      or str(ws.cell(header_row, c).value or "").strip().lower() == "shipped"),
                     None)
    if ship_col is None:
        return None
    prev = str(ws.cell(header_row, ship_col - 1).value or "").strip().lower() if ship_col > 1 else ""
    if "sop" in prev:
        return ship_col - 1
    _insert_col_keep_merges(ws, ship_col, 1)
    _copy_style(ws.cell(header_row, ship_col + 1), ws.cell(header_row, ship_col))
    ws.cell(header_row, ship_col).value = _SOP_HEADER
    from openpyxl.utils import get_column_letter
    ship_width = ws.column_dimensions[get_column_letter(ship_col + 1)].width
    if ship_width:
        ws.column_dimensions[get_column_letter(ship_col)].width = ship_width
    return ship_col


def _find_sheet(wb, *keyword_sets):
    """Find a sheet by fuzzy name. Each keyword_set is a tuple of words that must
    ALL appear in the sheet name (case-insensitive). Exact match wins first.
    Returns the worksheet, or None if no sheet matches — so a differently-shaped
    template simply skips tabs it doesn't have instead of crashing."""
    names = {name.strip().lower(): name for name in wb.sheetnames}
    for kws in keyword_sets:
        if isinstance(kws, str):
            kws = (kws,)
        joined = " ".join(kws)
        if joined in names:                      # exact
            return wb[names[joined]]
    for kws in keyword_sets:
        if isinstance(kws, str):
            kws = (kws,)
        for low, orig in names.items():
            if all(k in low for k in kws):
                return wb[orig]
    return None


# Third-stage names we've actually seen used in the wild. The real name is
# now discovered fresh per-report from the raw data (analysis.ReportData.
# third_stage_name) rather than hardcoded — this list exists only so an OLD
# sample template's own static text (baked in from whenever it was built)
# gets relabeled to the CURRENT name instead of staying stuck on a past one.
# A brand-new never-before-seen historical name in some template wouldn't be
# caught here; that's a cosmetic gap, not a data-correctness one (the actual
# attempt matching in analysis.py never depends on this list).
_KNOWN_LEGACY_THIRD_NAMES = {"navin", "blues"}


def _relabel_operator_aliases(wb, third_stage_name: str):
    """Rewrite any known-legacy third-stage name in a sheet's own static text
    (column headers, title bars) to the CURRENT report's actual third-stage
    name. Whole-word, case-insensitive; skips formula cells so we never touch
    a formula's text."""
    legacy = _KNOWN_LEGACY_THIRD_NAMES - {third_stage_name.strip().lower()}
    if not legacy:
        return
    pattern = re.compile(r'\b(' + '|'.join(re.escape(a) for a in legacy) + r')\b', re.IGNORECASE)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and not v.startswith("=") and pattern.search(v):
                    cell.value = pattern.sub(third_stage_name, v)


def fill_report(wb, R: ReportData, ots: Optional[OkToShip], log=lambda _m: None):
    # (sheet finder, fill function). Missing tabs are skipped with a note.
    tabs = [
        (("summary",),                    lambda ws: _fill_summary(ws, R, ots)),
        (("operator", "summary"),         lambda ws: _fill_operator_summary(ws, R)),
        (("daily",),                      lambda ws: _fill_daily_activity(ws, R)),
        (("device", "attempt"),           lambda ws: _fill_device_attempts(ws, R)),
        (("failed",),                     lambda ws: _fill_failed_products(ws, R)),
        (("device", "detail"),            lambda ws: _fill_device_detail(ws, R)),
        (("anomal",),                     lambda ws: _fill_anomalies(ws, R)),
        (("raw", "attempt"),              lambda ws: _fill_raw_attempts(ws, R, ots)),
    ]
    for keywords, fn in tabs:
        ws = _find_sheet(wb, keywords)
        if ws is None:
            log(f"  note: no tab matching {keywords} in this template — skipped")
            continue
        fn(ws)

    # The complete per-unit test history lives in "(Raw) Attempts", so the
    # separate "MC Retest Units" tab is removed from the output entirely.
    mc = _find_sheet(wb, ("retest",))
    if mc is not None and len(wb.sheetnames) > 1:
        wb.remove(mc)

    # The sample template's own static text (column headers, title bars like
    # "...OBA → Navin in strict order") is byte-copied and never touched by
    # the fill functions above, which only ever WRITE data — so if an
    # operator was renamed (e.g. Navin -> Blues) and this sample predates
    # that, its baked-in labels would still say the old name even though
    # every actual data cell (Tester columns, Operator Summary rows, the
    # narrative text we generate) already reflects the raw data correctly.
    # Sweep every sheet and relabel any such stale mention to the current
    # canonical name, so the output always matches who the raw data says
    # tested the units, regardless of which name the template was built with.
    _relabel_operator_aliases(wb, R.third_stage_name)


# --------------------------------------------------------------------------- #
FAIL_START = 11          # first failure-breakdown data row in the template
ANALYSIS_ROW = 17        # "Final Analysis" header row in the template


def _fill_summary(ws, R: ReportData, ots: Optional[OkToShip]):
    # Report on the CERTIFIED lot only (units whose final test passed).
    unique = R.final_pass
    _set(ws, 1, A, f"PO {R.po} — {R.product} Certification Report")
    _set(ws, 2, A, f"Product: {R.product}  |  PO: {R.po}  |  Build Code: {R.build_code}  |  "
                   f"CPN: {R.cpn}  |  Period: {R.period}  |  Operators: {R.third_stage_name} · OBA · Operator")

    # Some templates insert a "Shipment Qty" column right after Unique Devices,
    # pushing Final PASS/FAIL/Success Rate/Total Test Runs/Operators one column
    # right; detect it instead of assuming a fixed 6- or 7-column layout.
    header_row = _header_row_of(ws, ("unique device",), default=4)
    has_shipment_col = "shipment" in str(ws.cell(header_row, B).value or "").lower()
    s = 1 if has_shipment_col else 0
    from openpyxl.utils import get_column_letter
    pass_c, fail_c, rate_c, runs_c, ops_c = B + s, C + s, D + s, E + s, F + s

    _set(ws, 5, A, unique)
    if has_shipment_col:
        # Shipment Qty = units marked shipped (Shipped Status = TRUE in the
        # (Raw) Attempts tab), recomputed fresh every time the report is built.
        shipped_qty = sum(1 for d in R.all_devices
                          if (ots.shipped_status(d.imei) if ots else True))
        _set(ws, 5, B, shipped_qty)
    # Final PASS / Final FAIL / Total Test Runs are pulled live from the
    # Operator Summary tab's own totals (row 9 = certified-lot totals, row 16
    # = all-attempts totals) so the two tabs can never disagree.
    _set(ws, 5, pass_c, "='Operator Summary'!C9")
    _set(ws, 5, fail_c, "='Operator Summary'!D9")
    _set(ws, 5, rate_c, f"={get_column_letter(pass_c)}5/A5")   # Final PASS / Unique Devices
    _set(ws, 5, runs_c, "='Operator Summary'!B16")
    _set(ws, 5, ops_c, 3)

    # --- Failure breakdown (rows 11+) ---------------------------------- #
    nf = len(R.failure_breakdown)
    total_fail = sum(c for _, c in R.failure_breakdown) or 1
    # The "Final Analysis" banner's row varies a bit between template versions
    # (some reserve an extra blank buffer row before it), so detect its real
    # position on the PRISTINE template rather than assuming a fixed row —
    # otherwise the row-insert below fires one row early/late and the banner's
    # blue styling ends up on the wrong (narrative) line instead of the title.
    analysis_row = _header_row_of(ws, ("final analysis",), default=ANALYSIS_ROW,
                                  scan=ANALYSIS_ROW + 10)
    # keep one blank row before the analysis block; insert rows if we need more
    insert_count = max(0, (FAIL_START + nf + 1) - analysis_row)
    if insert_count:
        _insert_rows_keep_merges(ws, analysis_row, insert_count)
    # some templates merge spare failure rows full-width (e.g. A13:G13); unmerge
    # the rows we're about to fill so Occurrences/% (cols B/C) are writable
    for r in range(FAIL_START, FAIL_START + nf):
        _unmerge_region(ws, r, 1, 12)
    for idx, (reason, count) in enumerate(R.failure_breakdown):
        r = FAIL_START + idx
        _clone_row_style(ws, FAIL_START, r, 3)
        _set(ws, r, A, reason)
        _set(ws, r, B, count)
        _set(ws, r, C, count / total_fail)
    # clear any leftover template failure rows (when fewer reasons than before)
    for r in range(FAIL_START + nf, analysis_row):
        _clear_cells(ws, r, 3)

    # --- Final Analysis narrative (regenerated with real numbers) ------ #
    a0 = analysis_row + insert_count           # shifted "Final Analysis" header row
    _set(ws, a0, A, "Final Analysis")
    lines = _analysis_lines(R)
    for i in range(6):                          # template provides 6 narrative rows
        _set(ws, a0 + 1 + i, A, lines[i] if i < len(lines) else None)


def _analysis_lines(R: ReportData) -> list:
    tag = f"PO {R.po} {R.build_code}".strip()
    pass_fail = (f"*All {R.final_pass} certified units are PASS. "
                 f"There are 0 FAIL units for {tag}.")
    bar = ('*The bar chart in the "Daily Activity & Bar Chart" tab shows all PASS '
           "units per operator per date.")
    extra = f", {R.n_all3_unordered} had all three stages out of order" if R.n_all3_unordered else ""
    other = f", {R.n_other_path} took another path" if R.n_other_path else ""
    anomaly = (f"*Of the {len(R.anomalies)} anomalous units (not in strict "
               f"Operator→OBA→{R.third_stage_name} order): {R.n_oba_only} went Operator→OBA only, "
               f"{R.n_third_only} went Operator→{R.third_stage_name} only{extra}{other}. "
               f"{R.n_completed_ordered} units completed all 3 stages in order.")
    failed = ("*Since all certified units are PASS, there are no failed products "
              "to compare against the shipping list.")
    excluded = len(R.all_devices) - R.final_pass
    ship = ('*Shipped Status (in the "(Raw) Attempts" tab) is set from the OK-TO-SHIP '
            "Balance status" + (f"; {excluded} unit(s) did not reach certification and "
            "are listed only in the (Raw) Attempts tab." if excluded else "."))
    return [pass_fail, bar, anomaly, failed, ship]


def _fill_operator_summary(ws, R: ReportData):
    _set(ws, 1, A, f"PO {R.po} — {R.product} — Operator / Tester Summary")
    _set(ws, 2, A, f"{R.third_stage_name} · OBA · Operator  |  Station: {', '.join(R.stations)}  |  Build Code: {R.build_code}")
    order = [R.third_stage_name, "OBA", "operator"]
    # Section A — final cert responsibility (rows 6,7,8; total 9)
    for i, op in enumerate(order):
        r = 6 + i
        s = R.op_summary_final[op]
        _set(ws, r, A, "Operator" if op == "operator" else op)
        _set(ws, r, B, s["devices"])
        _set(ws, r, C, s["pass"])
        _set(ws, r, D, s["fail"])
        _set(ws, r, F, ", ".join(sorted(s["stations"])))
        _set(ws, r, G, ", ".join(sorted(s["months"])))
    _set(ws, 9, B, sum(R.op_summary_final[o]["devices"] for o in order))
    _set(ws, 9, C, sum(R.op_summary_final[o]["pass"] for o in order))
    _set(ws, 9, D, sum(R.op_summary_final[o]["fail"] for o in order))
    # Section B — all attempts (rows 13,14,15; total 16)
    for i, op in enumerate(order):
        r = 13 + i
        s = R.op_summary_all[op]
        _set(ws, r, A, "Operator" if op == "operator" else op)
        _set(ws, r, B, s["attempts"])
        _set(ws, r, C, s["pass"])
        _set(ws, r, D, s["fail"])
        _set(ws, r, F, ", ".join(sorted(s["stations"])))
        _set(ws, r, G, ", ".join(sorted(s["months"])))
    _set(ws, 16, B, sum(R.op_summary_all[o]["attempts"] for o in order))
    _set(ws, 16, C, sum(R.op_summary_all[o]["pass"] for o in order))
    _set(ws, 16, D, sum(R.op_summary_all[o]["fail"] for o in order))


def _capture_daily_highlights(ws, start, scan=60):
    """From the template rows, learn each operator column's fill for a value>0
    cell (highlight) and a value==0 cell (plain) — so we can highlight by VALUE."""
    hi = {}
    zero = {}
    for c in (B, C, D):
        for r in range(start, start + scan):
            v = ws.cell(r, c).value
            if not isinstance(v, (int, float)):
                continue
            fill = ws.cell(r, c).fill
            if v > 0 and c not in hi and fill and fill.patternType:
                hi[c] = copy(fill)
            elif v == 0 and c not in zero and fill and fill.patternType:
                zero[c] = copy(fill)
            if c in hi and c in zero:
                break
    return hi, zero


def _fill_daily_activity(ws, R: ReportData):
    _set(ws, 1, A, f"PO {R.po} — {R.product}")
    dates = sorted(R.daily_activity)
    start = 5
    tmpl = start
    hi, zero = _capture_daily_highlights(ws, start)   # value-based fills from template
    totals = {B: 0, C: 0, D: 0}
    for idx, dt in enumerate(dates):
        r = start + idx
        _clone_row_style(ws, tmpl, r, 5)
        cnt = R.daily_activity[dt]
        _set(ws, r, A, dt)
        vals = {B: cnt.get(R.third_stage_name, 0), C: cnt.get("OBA", 0), D: cnt.get("operator", 0)}
        for c, v in vals.items():
            _set(ws, r, c, v)
            fill = (hi.get(c) if v > 0 else zero.get(c))   # highlight only when > 0
            if fill is not None:
                ws.cell(r, c).fill = copy(fill)
            totals[c] += v
        _set(ws, r, E, f"=SUM(B{r}:D{r})")
    total_row = start + len(dates)
    _clone_row_style(ws, 9, total_row, 5)   # style of the template TOTAL row
    _set(ws, total_row, A, "TOTAL")
    last = total_row - 1
    _set(ws, total_row, B, f"=SUM(B{start}:B{last})")
    _set(ws, total_row, C, f"=SUM(C{start}:C{last})")
    _set(ws, total_row, D, f"=SUM(D{start}:D{last})")
    _set(ws, total_row, E, f"=SUM(E{start}:E{last})")
    # the TOTAL row's B/C/D are formulas, so highlight them from the actual
    # summed values (computed above) the same way the per-date rows are —
    # cloning the template's static TOTAL-row style alone left them unhighlighted
    # whenever a column's total was > 0 but its formatting wasn't already so.
    for c, v in totals.items():
        fill = hi.get(c) if v > 0 else zero.get(c)
        if fill is not None:
            ws.cell(total_row, c).fill = copy(fill)
    # clear any leftover template rows between new total and old total(9)
    for r in range(total_row + 1, 10):
        _clear_cells(ws, r, 5)
    title = f"Daily Activity — PO {R.po} {R.product} (Build {R.build_code})"
    _fix_daily_chart(ws, first=start, last=last, title=title)


def _set_chart_title(chart, title):
    """Replace the chart's title text, preserving its existing formatting."""
    try:
        runs = chart.title.tx.rich.p[0].r
        if runs:
            runs[0].t = title
            for extra in runs[1:]:
                extra.t = ""
            return
    except Exception:
        pass
    try:
        chart.title = title
    except Exception:
        pass


def _fix_daily_chart(ws, first, last, title=None):
    """Point the bar-chart series at exactly the date rows (excluding TOTAL)
    and refresh the chart title for the current build."""
    import re
    sheet = ws.title
    for chart in getattr(ws, "_charts", []):
        if title:
            _set_chart_title(chart, title)
        for s in chart.series:
            for ref in (getattr(s, "val", None), getattr(s, "cat", None)):
                nr = getattr(ref, "numRef", None) or getattr(ref, "strRef", None) if ref else None
                if nr and nr.f:
                    # keep the column letter(s), rewrite the row span
                    m = re.search(r"\$([A-Z]+)\$\d+:\$([A-Z]+)\$\d+", nr.f)
                    if m:
                        c1, c2 = m.group(1), m.group(2)
                        nr.f = f"'{sheet}'!${c1}${first}:${c2}${last}"


def _fill_device_attempts(ws, R: ReportData):
    # Detect the header row FIRST (on the pristine template) — some template
    # versions have a blank row 1 + title row 2 + header row 3, others put the
    # title directly in row 1 + header row 2, so title/data rows are derived
    # relative to the header rather than hardcoded.
    header_row = _header_row_of(ws, ("imei", "row label"), default=3)
    _set(ws, header_row - 1, A, f"Summary of ALL {len(R.certified_devices)} Units Tests")
    dev_col = _ensure_device_id_col(ws, header_row, A)
    # Device ID pushes everything after it one column right: op/oba/third
    # move from B/C/D to C/D/E, and approval moves from F to G.
    start = header_row + 1
    tmpl = start
    # Reuse the sample's own approval formula (captured before we overwrite
    # the template's first data row). Fall back to the known formula.
    approval_formula = _formula_at(ws, start, G) or \
        f'=IF((IF(C{start}>=1,1,0)+IF(D{start}>=1,1,0)+IF(E{start}>=1,1,0))>=2,"PASS","FAIL")'
    old_last = _detect_last(ws, A, start)
    for idx, d in enumerate(R.certified_devices):
        r = start + idx
        _clone_row_style(ws, tmpl, r, G)
        _set(ws, r, A, _imei_val(d.imei))
        _set(ws, r, dev_col, d.dev_id)
        _set(ws, r, C, d.op_attempts)
        _set(ws, r, D, d.oba_attempts)
        _set(ws, r, E, d.third_attempts)
        _propagate_formula(ws, approval_formula, start, r, G)
    for r in range(start + len(R.certified_devices), max(old_last, ws.max_row) + 1):
        _clear_cells(ws, r, G)


def _fill_failed_products(ws, R: ReportData):
    _set(ws, 1, A, f"Details: PO {R.po} {R.build_code} — All {len(R.certified_devices)} units are "
                   f"PASS. 0 failed units. Anomalous units (incomplete process) shown below.")
    start = 4
    tmpl = start
    old_last = _detect_last(ws, A, start)
    for idx, d in enumerate(R.anomalies):
        r = start + idx
        _clone_row_style(ws, tmpl, r, 2)
        _set(ws, r, A, d.imei)
        _set(ws, r, B, "No")
    for r in range(start + len(R.anomalies), max(old_last, ws.max_row) + 1):
        _clear_cells(ws, r, 2)


def _device_detail_value(header: str, d: Device, idx: int, third_name: str):
    """Map a Device Detail column HEADER to the right value for device d.
    Header-driven so the tab fills whatever columns a given template has
    (the sample may add/remove columns). Unknown headers return _SKIP."""
    h = header.strip().lower()
    if h in ("#", "no", "no.", "s/n", "sn no", "index"):
        return idx + 1
    if "imei" in h or "deveui" in h or (h in ("device", "device id", "serial", "serial number")):
        return _imei_val(d.imei)
    if "(dev" in h or "dev:" in h:
        return d.dev_id
    if h == "cpn" or "cpn" in h:
        return d.cpn
    if "certifying" in h or ("final" in h and "operator" in h):
        return "Operator" if d.final_operator == "operator" else d.final_operator
    if "first" in h and "date" in h:
        return d.first_run_date
    if "first" in h and "time" in h:
        return d.first_run_time
    if ("cert" in h or "final" in h) and "date" in h:
        return d.final_cert_date
    if ("cert" in h or "final" in h) and "time" in h:
        return d.final_cert_time
    if "all operator" in h or ("operators" in h and "attempt" not in h):
        return d.all_operators
    if "month" in h:
        return d.months
    if "station" in h:
        return ", ".join(d.stations)
    if "result" in h:
        return "PASS" if d.certified else "FAIL"
    if "operator" in h and "attempt" in h:
        return d.op_attempts
    if "oba" in h and "attempt" in h:
        return d.oba_attempts
    if (third_name.lower() in h or "navin" in h or "blues" in h) and "attempt" in h:
        return d.third_attempts
    if "total" in h and "attempt" in h:
        return d.total_attempts
    if "fail" in h and "reason" in h:
        return "; ".join(d.fail_reasons) if d.fail_reasons else None
    if "modem" in h or "firmware" in h:
        return d.modem
    return _SKIP


def _fill_device_detail(ws, R: ReportData):
    _set(ws, 1, A, f"PO {R.po} — {R.product} (Build {R.build_code}) — All {len(R.certified_devices)} Devices")
    _set(ws, 2, A, f"CPN: {R.cpn}  |  Ordering Code: {R.ordering_code}  |  Station: {', '.join(R.stations)}  |  "
                   f"Operators: {R.third_stage_name} · OBA · Operator  |  Build: {R.build_code} (PO {R.po}) {R.product}  |  "
                   f"PASS: {R.final_pass}  |  FAIL: 0")
    header_row, start = 3, 4
    imei_col = next((c for c in range(1, ws.max_column + 1)
                      if "imei" in str(ws.cell(header_row, c).value or "").lower()
                      or "deveui" in str(ws.cell(header_row, c).value or "").lower()), 2)
    _ensure_device_id_col(ws, header_row, imei_col)
    ncols = ws.max_column
    # read the column headers so we fill by MEANING, not fixed position
    headers = [(c, str(ws.cell(header_row, c).value or "")) for c in range(1, ncols + 1)]
    old_last = _detect_last(ws, imei_col, start)
    for idx, d in enumerate(R.certified_devices):
        r = start + idx
        _clone_row_style(ws, start, r, ncols)
        for c, h in headers:
            v = _device_detail_value(h, d, idx, R.third_stage_name)
            if v is not _SKIP:
                _set(ws, r, c, v)
    for r in range(start + len(R.certified_devices), max(old_last, ws.max_row) + 1):
        _clear_cells(ws, r, ncols)


def _fill_anomalies(ws, R: ReportData):
    header_row = _header_row_of(ws, ("imei",), default=3)
    dev_col = _ensure_device_id_col(ws, header_row, A)
    # Device ID pushes missing_stages/result from B/C to C/D.
    start = 4
    tmpl = start
    old_last = _detect_last(ws, A, start)
    for idx, d in enumerate(R.anomalies):
        r = start + idx
        _clone_row_style(ws, tmpl, r, D)
        _set(ws, r, A, d.imei)
        _set(ws, r, dev_col, d.dev_id)
        _set(ws, r, C, d.missing_stages)
        _set(ws, r, D, "PASS" if d.certified else "FAIL")
    for r in range(start + len(R.anomalies), max(old_last, ws.max_row) + 1):
        _clear_cells(ws, r, D)


def _add_true_false_highlight(ws, col, first_row, last_row):
    """Apply Excel's standard 'Highlight Cells Rules > Text that Contains'
    green/red preset to a TRUE/FALSE column (matches the colours the user
    configured by hand: green fill+font for TRUE, red for FALSE)."""
    if last_row < first_row:
        return
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    col_letter = get_column_letter(col)
    rng = f"{col_letter}{first_row}:{col_letter}{last_row}"
    anchor = f"{col_letter}{first_row}"
    green = DifferentialStyle(font=Font(color="FF006100"),
                              fill=PatternFill(bgColor="FFC6EFCE"))
    red = DifferentialStyle(font=Font(color="FF9C0006"),
                            fill=PatternFill(bgColor="FFFFC7CE"))
    ws.conditional_formatting.add(rng, Rule(
        type="containsText", operator="containsText", text="FALSE", dxf=red,
        formula=[f'NOT(ISERROR(SEARCH("FALSE",{anchor})))']))
    ws.conditional_formatting.add(rng, Rule(
        type="containsText", operator="containsText", text="TRUE", dxf=green,
        formula=[f'NOT(ISERROR(SEARCH("TRUE",{anchor})))']))


# SOP Followed: Excel's standard "Highlight Cells" preset colors (same green
# /red already used for Shipped Status), plus the matching yellow preset for
# the "retested but still not resolved" case. Direct per-cell fills, not
# conditional formatting — "No" needs two different colors (yellow vs red)
# depending on severity, which plain text-content matching can't tell apart.
_SOP_STYLE = {
    "yes":    (PatternFill("solid", fgColor="FFC6EFCE"), Font(color="FF006100"), "Yes"),
    "yellow": (PatternFill("solid", fgColor="FFFFEB9C"), Font(color="FF9C6500"), "No"),
    "red":    (PatternFill("solid", fgColor="FFFFC7CE"), Font(color="FF9C0006"), "No"),
}
_SOP_IMEI_FILL = {"yellow": PatternFill("solid", fgColor="FFFFEB9C"),
                  "red": PatternFill("solid", fgColor="FFFFC7CE")}


def _fill_raw_attempts(ws, R: ReportData, ots: Optional[OkToShip]):
    header_row = _header_row_of(ws, ("imei",), default=2)
    dev_col = _ensure_device_id_col(ws, header_row, A)
    sop_col = _ensure_sop_column(ws, header_row)
    start = 3
    summary_cols = F + (1 if dev_col else 0) + (1 if sop_col else 0)
    _write_attempt_table(ws, R.all_devices, RAW_ATTEMPTS_BLOCKS, ots, start=start,
                         summary_cols=summary_cols, dev_col=dev_col, sop_col=sop_col)
    ship_col = summary_cols
    _add_true_false_highlight(ws, ship_col, start, start + len(R.all_devices) - 1)


def _fill_mc_retest(ws, R: ReportData, ots: Optional[OkToShip]):
    n = _write_attempt_table(ws, R.mc_retest, MC_BLOCKS, ots, start=3,
                             clear_to=7)  # template used rows 3..7
    # Fix the COUNTIF ranges in the distribution block to cover the data rows.
    last = 3 + max(n, 1) - 1
    for r in range(10, 16):          # 1..6 attempts rows
        attempts = r - 9
        _set(ws, r, B, f"=COUNTIF($B$3:$B${last},{attempts})")
        _set(ws, r, C, f"=COUNTIF($C$3:$C${last},{attempts})")
        _set(ws, r, D, f"=COUNTIF($D$3:$D${last},{attempts})")
    _set(ws, 17, B, f"=SUM(B10:B15)")
    _set(ws, 17, C, f"=SUM(C10:C15)")
    _set(ws, 17, D, f"=SUM(D10:D15)")


def _count_blocks(ws, block_start=7):
    """How many 'Test N' column blocks the template actually has (each block is
    BLOCK_COLS wide, starting at `block_start`). Detected from the row-2
    sub-headers so we never assume a fixed count."""
    n, c = 0, block_start
    while c <= ws.max_column and ws.cell(2, c).value not in (None, ""):
        n += 1
        c += BLOCK_COLS
    return n


def _unmerge_region(ws, row, col_start, col_end):
    """Unmerge any merged range overlapping the given row/column span."""
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col_end and mr.max_col >= col_start:
            ws.unmerge_cells(str(mr))


def _clear_block_headers(ws, from_block, to_block, block_start=7):
    """Remove the header + sub-header + data-cell styling for blocks the report
    doesn't need (template had more Test blocks than this build uses)."""
    for bi in range(from_block, to_block):
        base = block_start + bi * BLOCK_COLS
        _unmerge_region(ws, 1, base, base + BLOCK_COLS - 1)
        for rr in (1, 2):
            for k in range(BLOCK_COLS):
                cell = ws.cell(rr, base + k)
                cell.value = None
                cell.fill = _BLANK_FILL
                cell.border = _BLANK_BORDER
                cell.font = _DEFAULT_FONT


def _ensure_blocks(ws, have_blocks, need_blocks, block_start=7):
    """Add extra 'Test N' column blocks (header + sub-headers + widths) so the
    tab can hold devices with more attempts than the template provided."""
    if need_blocks <= have_blocks:
        return
    from openpyxl.utils import get_column_letter
    src = block_start  # Test-1 block starts here
    for bi in range(have_blocks, need_blocks):
        base = block_start + bi * BLOCK_COLS
        # ensure the target region is not part of a leftover merged range
        _unmerge_region(ws, 1, base, base + BLOCK_COLS - 1)
        # row 1: group header "Test N" (merged across the 6 columns)
        for k in range(BLOCK_COLS):
            _copy_style(ws.cell(1, src + k), ws.cell(1, base + k))
            ws.cell(1, base + k).value = None
        ws.cell(1, base).value = f"Test {bi + 1}"
        try:
            ws.merge_cells(start_row=1, start_column=base,
                           end_row=1, end_column=base + BLOCK_COLS - 1)
        except Exception:
            pass
        # row 2: column sub-headers (Tester, Station, Date, Time, Result, Fail Reason)
        for k in range(BLOCK_COLS):
            s = ws.cell(2, src + k)
            t = ws.cell(2, base + k)
            _copy_style(s, t)
            t.value = s.value
        # match column widths to the Test-1 block
        for k in range(BLOCK_COLS):
            w = ws.column_dimensions[get_column_letter(src + k)].width
            ws.column_dimensions[get_column_letter(base + k)].width = w


def _write_attempt_table(ws, devices, template_blocks, ots, start, clear_to=None,
                         summary_cols=6, dev_col=None, sop_col=None):
    """Shared writer for the (Raw) Attempts / MC Retest 'Device Summary + Test N'
    layout. Grows the number of Test blocks to fit the device with the most
    attempts, and gives every block the same formatting as the Test-1 block.

    `summary_cols` is how many summary columns precede the Test-N blocks: 6
    normally (IMEI/op/oba/third/total/shipped), +1 when a Device ID column
    has been inserted after IMEI, +1 more when an SOP Followed column has
    been inserted before Shipped Status."""
    tmpl = start
    block_start = summary_cols + 1
    max_attempts = max((d.total_attempts for d in devices), default=0)
    # detect how many Test blocks the template really has (don't assume a count)
    real_blocks = max(_count_blocks(ws, block_start), 1)
    # the report shows exactly as many Test blocks as the data needs
    n_blocks = max(max_attempts, 1)
    # capture PASS / FAIL / empty block styles BEFORE we overwrite template rows
    pass_s, fail_s, empty_s = _capture_block_styles(ws, start, block_start)
    if n_blocks > real_blocks:
        _ensure_blocks(ws, real_blocks, n_blocks, block_start)        # add columns for extra attempts
    elif n_blocks < real_blocks:
        _clear_block_headers(ws, n_blocks, real_blocks, block_start)  # drop unused template blocks
    process_blocks = max(n_blocks, real_blocks)          # also blank any dropped-block data cells
    max_col = summary_cols + process_blocks * BLOCK_COLS
    old_last = _detect_last(ws, A, start)
    if clear_to is not None:
        old_last = max(old_last, clear_to)

    def _style_block(r, base, snaps):
        for k in range(BLOCK_COLS):
            if snaps and snaps[k] is not None:
                _apply_snap(ws.cell(r, base + k), snaps[k])
            else:
                _copy_style(ws.cell(tmpl, block_start + k), ws.cell(r, base + k))

    sum_snap = _capture_row(ws, tmpl, summary_cols)   # summary cols, captured once
    tmpl_height = ws.row_dimensions[tmpl].height
    # with a Device ID column, op/oba/third/total move one column right; SOP
    # Followed (if present) sits right after Total, pushing Shipped Status
    # one column further right still.
    s = 1 if dev_col else 0
    s2 = 1 if sop_col else 0
    op_c, oba_c, third_c, tot_c = B + s, C + s, D + s, E + s
    ship_c = F + s + s2
    for idx, d in enumerate(devices):
        r = start + idx
        # summary columns keep the template row's style (shared, no copy)
        _apply_row(ws, r, sum_snap)
        ws.row_dimensions[r].height = tmpl_height
        _set(ws, r, A, _imei_val(d.imei))
        if dev_col:
            _set(ws, r, dev_col, d.dev_id)
        _set(ws, r, op_c, d.op_attempts)
        _set(ws, r, oba_c, d.oba_attempts)
        _set(ws, r, third_c, d.third_attempts)
        _set(ws, r, tot_c, d.total_attempts)
        if sop_col:
            fill, font, text = _SOP_STYLE[d.sop_status]
            _set(ws, r, sop_col, text)
            ws.cell(r, sop_col).fill = fill
            ws.cell(r, sop_col).font = font
            if d.sop_status in _SOP_IMEI_FILL:
                ws.cell(r, A).fill = _SOP_IMEI_FILL[d.sop_status]
            elif dev_col:
                # Not a violation: make sure column A's background reflects
                # this row's actual zebra banding rather than whatever
                # `sum_snap` happened to capture for column A — which, if
                # this template is itself a PREVIOUS run's output reused as
                # the next sample, could still be a violation color left
                # over from a different device that used to occupy this row.
                # Device ID never gets special per-row coloring, so its fill
                # for this same row is a reliable "what should this row's
                # banding actually look like" reference.
                ws.cell(r, A).fill = copy(ws.cell(r, dev_col).fill)
        _set(ws, r, ship_c, ots.shipped_status(d.imei) if ots else True)
        for bi in range(process_blocks):
            base = block_start + bi * BLOCK_COLS
            if bi < len(d.events):
                e = d.events[bi]
                # PASS-green / FAIL-red styling per attempt result
                _style_block(r, base, (pass_s if e.result == "PASS" else fail_s) or pass_s)
                _set(ws, r, base + 0, "Operator" if e.tester == "operator" else e.tester)
                _set(ws, r, base + 1, e.station)
                _set(ws, r, base + 2, e.date)
                _set(ws, r, base + 3, e.time)
                _set(ws, r, base + 4, e.result)
                # a failed attempt must always carry a reason
                reason = e.fail_reason or ("Unspecified failure" if e.result == "FAIL" else None)
                _set(ws, r, base + 5, reason)
            else:
                # empty block WITHIN a data row — keep the sample's row/block
                # styling (so the row looks uniform), just clear the value
                _style_block(r, base, empty_s)
                for k in range(BLOCK_COLS):
                    _set(ws, r, base + k, None)
    # below the last data row: strip ALL leftover template formatting/banding,
    # right down to the template's formatted extent (not just its last value row)
    clear_end = max(old_last, ws.max_row)
    for r in range(start + len(devices), clear_end + 1):
        _clear_cells(ws, r, max_col)
    return len(devices)


# --------------------------------------------------------------------------- #
def _detect_last(ws, col, start, guard=20000):
    """Return the last consecutive non-empty data row starting at `start`."""
    r = start
    last = start - 1
    blanks = 0
    while r < start + guard:
        if ws.cell(r, col).value not in (None, ""):
            last = r
            blanks = 0
        else:
            blanks += 1
            if blanks > 3:
                break
        r += 1
    return last


def _imei_val(imei: str):
    """Keep IMEIs as text to avoid Excel turning long numbers into 8.6242E+14."""
    return str(imei)
