"""
Parse the raw-data folder of per-device test-log files.

Each file is named by a device id (IMEI or DevEUI) and contains ONE OR MORE
JSON objects concatenated together (whitespace/newline separated) — one object
per test *attempt*. We stream-decode all objects in each file.

Every attempt object carries fields like:
  imei / deveui / device, build, operator, station, when (unix ts),
  status, fail_reason / fail / fail_test, cpn, ordering_code, sku, product
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional


_DECODER = json.JSONDecoder()


@dataclass
class Attempt:
    device_id: str
    operator: str
    station: str
    when: Optional[int]          # unix timestamp (seconds)
    status: str
    passed: bool                 # True if this attempt has no failure recorded
    fail_reason: str
    build_raw: str               # raw build string, e.g. "Build NA (PO 1199) NOTE-MBNAN v2.2"
    po: str                      # parsed PO number, e.g. "1199"
    build_code: str              # parsed 2-letter code, e.g. "NA"
    product: str                 # parsed product, e.g. "NOTE-MBNAN"
    cpn: str
    ordering_code: str
    sku: str
    source_file: str
    modem: str = ""               # modem firmware string, e.g. "EG915QNALGR01A04M04_..."
    dev_id: str = ""              # hardware UID from the "device":"dev:..." field


# --------------------------------------------------------------------------- #
# build-string parsing
# --------------------------------------------------------------------------- #
_PO_RE = re.compile(r"PO\s*[:#]?\s*(\d{3,5})", re.IGNORECASE)
_CODE_RE = re.compile(r"Build\s+([A-Z]{1,3})\b")
_PRODUCT_RE = re.compile(r"\b(NOTE-[A-Z0-9]+|STARNOTE[A-Z0-9\- .]*|TRACKNOTE[A-Z0-9\- .]*)", re.IGNORECASE)


def parse_build_string(build: str) -> tuple[str, str, str]:
    """Return (po, build_code, product) parsed from a raw build string.
    Tolerates short forms like "NOTE-MBNAN" (no PO / code)."""
    build = build or ""
    po = ""
    code = ""
    product = ""
    m = _PO_RE.search(build)
    if m:
        po = m.group(1)
    m = _CODE_RE.search(build)
    if m:
        code = m.group(1)
    m = _PRODUCT_RE.search(build)
    if m:
        product = m.group(1).upper().strip()
    return po, code, product


def _fail_of(obj: dict) -> str:
    """Return a failure message if this attempt failed, else "".
    Failures can be recorded under any of these fields (checked in priority
    order); `err` is included so modem/hardware errors aren't missed."""
    for k in ("fail_reason", "fail", "fail_test", "err"):
        v = obj.get(k)
        if v:
            return str(v)
    return ""


def _device_id_of(obj: dict, fallback: str) -> str:
    for k in ("imei", "deveui"):
        v = obj.get(k)
        if v:
            return str(v)
    dev = obj.get("device", "")
    if isinstance(dev, str) and dev.startswith("dev:"):
        return dev[4:]
    return fallback


def _dev_id_of(obj: dict) -> str:
    """Hardware UID reported under the "device" field, e.g. "dev:869830..."
    -> "869830...". This is distinct from the IMEI/DevEUI and is always
    present, even for early failures that report no IMEI yet."""
    dev = obj.get("device", "")
    if isinstance(dev, str) and dev.startswith("dev:"):
        return dev[4:]
    return str(dev) if dev else ""


def iter_json_objects(text: str) -> Iterator[dict]:
    """Yield each JSON object in a string that may hold several concatenated."""
    idx, n = 0, len(text)
    while idx < n:
        while idx < n and text[idx] in " \n\r\t":
            idx += 1
        if idx >= n:
            break
        try:
            obj, end = _DECODER.raw_decode(text, idx)
        except json.JSONDecodeError:
            break  # give up on the remainder of a malformed file
        if isinstance(obj, dict):
            yield obj
        idx = end


# Extensions we never try to parse as device JSON (avoids wasted reads).
_SKIP_EXT = {".xlsx", ".xls", ".xlsm", ".pdf", ".png", ".jpg", ".jpeg",
             ".gif", ".doc", ".docx", ".csv", ".mov", ".mp4", ".key"}


def parse_content(text: str, source_name: str, fallback_id: str) -> list[Attempt]:
    """Parse one file's/entry's text (one or many concatenated JSON objects).

    Each file holds exactly one physical device's attempts, and the file is named
    after that device. A unit's *reported* id can change across attempts (early
    failures report the hardware UID like 'dev:2030...' before the IMEI is read),
    so we key every attempt to the FILE's id (fallback_id) — not the per-object
    id — to keep one unit as one device."""
    attempts: list[Attempt] = []
    device_id = fallback_id                    # canonical id = the file name
    for obj in iter_json_objects(text):
        build_raw = str(obj.get("build", ""))
        po, code, product = parse_build_string(build_raw)
        fail = _fail_of(obj)
        attempts.append(Attempt(
            device_id=device_id,
            operator=str(obj.get("operator", "")),
            station=str(obj.get("station", "")),
            when=obj.get("when") if isinstance(obj.get("when"), (int, float)) else None,
            status=str(obj.get("status", "")),
            passed=(fail == ""),
            fail_reason=fail,
            build_raw=build_raw,
            po=po,
            build_code=code,
            product=product,
            cpn=str(obj.get("cpn", "")),
            ordering_code=str(obj.get("ordering_code", "")),
            sku=str(obj.get("sku", "")),
            source_file=source_name,
            modem=str(obj.get("modem", "")),
            dev_id=_dev_id_of(obj),
        ))
    return attempts


def parse_file(path: Path) -> list[Attempt]:
    try:
        text = path.read_text(errors="replace")
    except OSError:
        return []
    return parse_content(text, path.name, path.stem)


def _entry_ok(name: str) -> bool:
    if name.endswith("/") or "__MACOSX" in name:
        return False
    base = name.rsplit("/", 1)[-1]
    return bool(base) and not base.startswith(".")


def parse_zip(path: Path) -> list[Attempt]:
    """Read every device entry inside a .zip archive."""
    import zipfile
    out: list[Attempt] = []
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if not _entry_ok(name):
                    continue
                base = name.rsplit("/", 1)[-1]
                ext = ("." + base.rsplit(".", 1)[-1].lower()) if "." in base else ""
                if ext in _SKIP_EXT or ext == ".zip":
                    continue
                try:
                    text = z.read(name).decode("utf-8", "replace")
                except Exception:
                    continue
                stem = base.rsplit(".", 1)[0] if "." in base else base
                out.extend(parse_content(text, f"{path.name}:{name}", stem))
    except Exception:
        pass
    return out


def parse_raw_dir(raw_dir: Path) -> list[Attempt]:
    """Parse every file under raw_dir — loose files AND the contents of any .zip
    archives (recursing through sub-folders). Hidden / macOS junk is skipped."""
    raw_dir = Path(raw_dir)
    out: list[Attempt] = []
    if not raw_dir.exists():
        return out
    # allow passing a single file / zip directly
    if raw_dir.is_file():
        return parse_zip(raw_dir) if raw_dir.suffix.lower() == ".zip" else parse_file(raw_dir)
    for p in sorted(raw_dir.rglob("*")):
        if not p.is_file() or p.name.startswith(".") or "__MACOSX" in p.parts:
            continue
        ext = p.suffix.lower()
        if ext == ".zip":
            out.extend(parse_zip(p))
        elif ext not in _SKIP_EXT:
            out.extend(parse_file(p))
    return out
