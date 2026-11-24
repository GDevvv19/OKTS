"""
Report engine: orchestrates parsing the raw test logs + OK-TO-SHIP report, then
produces a new report by copying the sample template and filling it in.
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import openpyxl

from raw_parser import parse_raw_dir
from analysis import build_report_data, available_builds
from ok_to_ship import parse_ok_to_ship
from filler import fill_report


@dataclass
class ReportRequest:
    sample_path: Path            # template report whose format/formulas we clone
    raw_reports_dir: Path        # folder of unique-ID-coded test reports
    output_path: Path            # where to save the generated report
    po_number: str = ""
    build_code: str = ""
    ok_to_ship_path: Optional[Path] = None   # optional


def generate_report(req: ReportRequest, progress: Callable[[str], None] = lambda _: None) -> Path:
    progress("Validating inputs...")
    if not req.sample_path or not Path(req.sample_path).exists():
        raise FileNotFoundError(f"Sample template not found: {req.sample_path}")
    if not req.raw_reports_dir or not Path(req.raw_reports_dir).exists():
        raise FileNotFoundError(f"Raw-data folder not found: {req.raw_reports_dir}")
    if not (req.po_number or req.build_code):
        raise ValueError("Please enter a PO Number and/or Build Code to select the build.")

    progress("Reading raw test reports...")
    attempts = parse_raw_dir(Path(req.raw_reports_dir))
    progress(f"  parsed {len(attempts):,} test attempts")

    if not attempts:
        raise ValueError(
            "No readable test files were found in the raw-data folder.\n"
            "  • If the folder contains a .zip, unzip it first and point to the unzipped folder.\n"
            "  • Make sure you selected the folder that directly contains the device files.\n"
            f"  Folder: {req.raw_reports_dir}")

    progress("Aggregating device data...")
    R = build_report_data(attempts, po=req.po_number, build_code=req.build_code)
    if not R.all_devices:
        found = available_builds(attempts)
        lines = "\n".join(f"    PO {po}  |  Build Code {bc}  |  {prod}  ({n} records)"
                          for (po, bc, prod), n in found[:15])
        raise ValueError(
            f"No devices matched PO '{req.po_number}' / Build Code '{req.build_code}'.\n"
            f"These are the PO / Build Code combinations found in this folder:\n{lines}\n"
            "  Tip: leave a box blank to match on the other one only "
            "(a '—' above means that field is missing from the files).")
    progress("Aggregating device data...")
    progress(f"  {len(R.all_devices)} devices ({R.final_pass} certified, "
             f"{R.total_test_runs} test runs)")

    ots = None
    if req.ok_to_ship_path and Path(req.ok_to_ship_path).exists():
        progress("Reading OK-TO-SHIP report...")
        ots = parse_ok_to_ship(Path(req.ok_to_ship_path))
        held = len(ots.balance_imeis)
        progress(f"  {held} unit(s) held back (Balance status)")

    progress("Cloning sample template (formulas + charts preserved)...")
    out = Path(req.output_path)
    # Ensure a valid Excel extension even if the user typed a name without one
    # (e.g. "Report v1.4" would otherwise be read as a ".4" file and rejected).
    if out.suffix.lower() not in (".xlsx", ".xlsm"):
        out = out.with_name(out.name + ".xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(req.sample_path, out)

    progress("Filling report...")
    wb = openpyxl.load_workbook(out)
    try:
        # Sample templates are often a PREVIOUS generated report reused as the
        # next one's starting point. If that prior report's chart was ever
        # copy/pasted from yet another workbook, Excel can leave behind a
        # dangling "external workbook" link (xl/externalLinks/...) — a stale,
        # unused reference that nothing in the file actually depends on, but
        # which openpyxl's load+save round-trip does not reliably preserve.
        # Left in place, this is what makes Excel report "a problem with some
        # content" and offer to repair the file. Stripping it here both fixes
        # this output AND stops the artifact from propagating into whatever
        # future report reuses this one as ITS sample template.
        if wb._external_links:
            progress(f"  removing {len(wb._external_links)} stale external "
                     "workbook link(s) inherited from the sample (prevents "
                     "Excel 'repair file' prompts)")
            wb._external_links = []
        fill_report(wb, R, ots, log=progress)
        wb.save(out)
    finally:
        wb.close()

    progress(f"Done → {out}")
    return out
