"""
Look up a single device by its ID (IMEI / DevEUI) in the raw-data folder.

Because each raw file is *named* by the device id, we can find one device
without scanning the whole folder: try the exact filename first, then fall back
to a filename/field search.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

import zipfile

from raw_parser import Attempt, parse_content, parse_file, parse_raw_dir

# Match the report's convention: raw "when" is UTC, displayed in US Eastern.
_EASTERN = ZoneInfo("America/New_York")


@dataclass
class DeviceLookup:
    device_id: str
    attempts: list[Attempt]
    source_file: str

    @property
    def po(self) -> str:
        return next((a.po for a in self.attempts if a.po), "")

    @property
    def build_code(self) -> str:
        return next((a.build_code for a in self.attempts if a.build_code), "")

    @property
    def product(self) -> str:
        return next((a.product for a in self.attempts if a.product), "")

    @property
    def final_result(self) -> str:
        ev = self._sorted()
        return ("PASS" if ev[-1].passed else "FAIL") if ev else "—"

    def _sorted(self) -> list[Attempt]:
        return sorted(self.attempts, key=lambda a: (a.when or 0))


def _fmt(when) -> str:
    if not when:
        return "—"
    dt = datetime.datetime.fromtimestamp(when, tz=datetime.timezone.utc).astimezone(_EASTERN)
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def lookup_device(raw_dir, device_id: str,
                  po: str = "", build_code: str = "") -> Optional[DeviceLookup]:
    raw_dir = Path(raw_dir)
    device_id = (device_id or "").strip()
    if not device_id or not raw_dir.exists():
        return None

    attempts: list[Attempt] = []
    source = ""

    # 1) fast path — a loose file literally named after the device id
    for p in raw_dir.rglob(device_id):
        if p.is_file() and "__MACOSX" not in p.parts:
            attempts = parse_file(p)
            source = str(p)
            break

    # 1b) fast path — an entry named after the device id inside any .zip
    if not attempts:
        for zp in raw_dir.rglob("*.zip"):
            try:
                with zipfile.ZipFile(zp) as z:
                    for name in z.namelist():
                        if name.endswith("/") or "__MACOSX" in name:
                            continue
                        if name.rsplit("/", 1)[-1] == device_id:
                            text = z.read(name).decode("utf-8", "replace")
                            attempts = parse_content(text, f"{zp.name}:{name}", device_id)
                            source = f"{zp.name}:{name}"
                            break
            except Exception:
                pass
            if attempts:
                break

    # 2) fall back — scan everything, match on parsed device_id
    if not attempts:
        for a in parse_raw_dir(raw_dir):
            if a.device_id == device_id:
                attempts.append(a)
                source = source or a.source_file

    if not attempts:
        return None

    # optional PO / Build Code narrowing
    po = (po or "").strip()
    bc = (build_code or "").strip().upper()
    if po:
        attempts = [a for a in attempts if a.po == po] or attempts
    if bc:
        attempts = [a for a in attempts if a.build_code.upper() == bc] or attempts

    return DeviceLookup(device_id=device_id, attempts=attempts, source_file=source)


def format_lookup(dl: Optional[DeviceLookup], device_id: str) -> str:
    if dl is None:
        return f"Device '{device_id}' was not found in the raw-data folder."

    ev = dl._sorted()
    lines = [
        f"Device: {dl.device_id}",
        f"Product: {dl.product or '—'}    PO: {dl.po or '—'}    Build Code: {dl.build_code or '—'}",
        f"Final result: {dl.final_result}    Total attempts: {len(ev)}",
        f"Source file: {dl.source_file}",
        "",
        f"{'#':<3}{'Operator':<12}{'Station':<12}{'When (ET)':<22}{'Result':<7}Fail reason",
        "-" * 78,
    ]
    for i, a in enumerate(ev, 1):
        lines.append(
            f"{i:<3}{a.operator:<12}{a.station:<12}{_fmt(a.when):<22}"
            f"{('PASS' if a.passed else 'FAIL'):<7}{a.fail_reason}"
        )
    return "\n".join(lines)
