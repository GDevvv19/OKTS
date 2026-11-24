"""
Turn a flat list of raw Attempts into the structured data that populates every
tab of the certification report.

Rules confirmed against the PO 1218 sample:
  * Timestamps (`when`) are raw Unix seconds (UTC) and are rendered in US
    Eastern Time (America/New_York, DST-aware) throughout the report.
  * There are always exactly 3 test stages: "operator", "OBA", and a third
    (final-certifying) stage. The first two are fixed role codes; the third
    is whoever the raw data says it is — discovered fresh from each report's
    OWN attempts rather than matched against a hardcoded name, so the
    program never needs a code change when that role's real name changes
    (it has twice already: Navin, then Blues).
  * Devices are filtered by the user-entered PO number and Build Code.
  * A device is "certified" if its final (latest) attempt passed.
  * Shipped Status comes from the OK-TO-SHIP report (Balance status = held back);
    it is applied later, in the engine, not here.
  * SOP Followed (added 2026-11-05, refined 2026-11-06): checked per-tester.
    After a FAIL, that same tester must retest TWICE more and BOTH must PASS
    (checked against the two attempts right after the LAST fail) — else
    "No"/yellow. 3+ consecutive FAILs at one tester is "No"/red (checked
    regardless of a later recovery). Otherwise "Yes"/green.
"""

from __future__ import annotations

import collections
import datetime
from dataclasses import dataclass, field
from typing import Optional
from zoneinfo import ZoneInfo

from raw_parser import Attempt

# Used only when a build's data has zero third-stage attempts at all (so
# there's nothing to detect a name from) — a harmless placeholder, never
# what drives real report data.
_DEFAULT_THIRD_NAME = "Blues"

# All report timestamps are shown in US Eastern Time, not the raw UTC the
# stations log in — DST-aware (EST/EDT switches automatically by date).
_EASTERN = ZoneInfo("America/New_York")


def _detect_third_stage(attempts) -> str:
    """Find the real name of the third test stage directly from THIS build's
    own raw data: whichever non-"operator"/"OBA" operator string appears most
    often. Matching is case-insensitive (so "Blues"/"blues"/"BLUES" count as
    the same tester) but the most common exact spelling is kept for display,
    so the report always shows the name literally as the test records do."""
    counts: dict[str, int] = {}
    display: dict[str, str] = {}
    for a in attempts:
        raw = (a.operator or "").strip()
        low = raw.lower()
        if not raw or low in ("operator", "oba"):
            continue
        counts[low] = counts.get(low, 0) + 1
        if low not in display:
            display[low] = raw
    if not counts:
        return ""
    winner = max(counts, key=counts.get)
    return display[winner]


def _to_eastern(when: int) -> datetime.datetime:
    return datetime.datetime.fromtimestamp(when, tz=datetime.timezone.utc).astimezone(_EASTERN)


def _fmt_date(when: Optional[int]) -> str:
    if not when:
        return ""
    return _to_eastern(when).strftime("%Y-%m-%d")


def _fmt_time(when: Optional[int]) -> str:
    if not when:
        return ""
    return _to_eastern(when).strftime("%H:%M:%S")


@dataclass
class TestEvent:
    tester: str          # canonical operator
    station: str
    date: str
    time: str
    result: str          # "PASS" / "FAIL"
    fail_reason: str
    when: Optional[int]
    dev_id: str = ""     # hardware UID this specific attempt reported


@dataclass
class Device:
    imei: str
    events: list[TestEvent] = field(default_factory=list)
    cpn: str = ""
    ordering_code: str = ""
    modem: str = ""
    dev_id: str = ""      # hardware UID from the raw "device":"dev:..." field

    # derived
    stations: list = field(default_factory=list)      # distinct stations, in order
    fail_reasons: list = field(default_factory=list)  # distinct failure reasons
    op_attempts: int = 0
    oba_attempts: int = 0
    third_attempts: int = 0   # attempts by whoever the third (final-cert) stage is
    total_attempts: int = 0
    first_run_date: str = ""
    first_run_time: str = ""
    final_cert_date: str = ""
    final_cert_time: str = ""
    final_operator: str = ""
    all_operators: str = ""
    months: str = ""
    certified: bool = False          # final attempt passed
    completed_all_stages: bool = False   # went through operator -> OBA -> third stage, in order
    missing_stages: str = ""
    sop_status: str = "yes"           # "yes" / "yellow" / "red" — see finalize()

    def finalize(self, operators: list[str]):
        """`operators` is this report's own [operator, OBA, third-stage-name]
        list, discovered from the data rather than hardcoded."""
        ev = sorted(self.events, key=lambda e: (e.when or 0))
        self.events = ev
        third = operators[2] if len(operators) > 2 else None
        self.total_attempts = len(ev)
        self.op_attempts = sum(1 for e in ev if e.tester == "operator")
        self.oba_attempts = sum(1 for e in ev if e.tester == "OBA")
        self.third_attempts = sum(1 for e in ev if e.tester == third)
        if ev:
            self.first_run_date, self.first_run_time = ev[0].date, ev[0].time
            self.final_cert_date, self.final_cert_time = ev[-1].date, ev[-1].time
            self.final_operator = ev[-1].tester
            self.certified = ev[-1].result == "PASS"
            seen = []
            for e in ev:
                if e.tester not in seen:
                    seen.append(e.tester)
            self.all_operators = ", ".join(seen)
            self.months = ", ".join(sorted({e.date[:7] for e in ev if e.date}))
            # Device ID: prefer whatever a PASSING attempt reported (the
            # settled identifier); only if the unit never passed at all do we
            # fall back to whatever raw value is available (e.g. the early
            # hardware boot UID reported by a FAIL attempt).
            self.dev_id = (next((e.dev_id for e in ev if e.result == "PASS" and e.dev_id), "")
                           or next((e.dev_id for e in ev if e.dev_id), ""))
        # distinct stations (in first-seen order) and distinct failure reasons
        st_seen, self.stations = set(), []
        for e in ev:
            if e.station and e.station not in st_seen:
                st_seen.add(e.station)
                self.stations.append(e.station)
        fr_seen, self.fail_reasons = set(), []
        for e in ev:
            if e.result == "FAIL" and e.fail_reason and e.fail_reason not in fr_seen:
                fr_seen.add(e.fail_reason)
                self.fail_reasons.append(e.fail_reason)
        # strict-order completion: operator then OBA then the third stage, as a subsequence
        order = [e.tester for e in ev]
        self.completed_all_stages = _is_ordered_subsequence(order, operators)
        missing = [s for s in operators if s not in order]
        self.missing_stages = ", ".join(missing)
        # SOP compliance, checked independently per tester: after a FAIL, that
        # SAME tester must retest the unit TWO more times, and BOTH retests
        # must PASS — the two attempts immediately following the tester's
        # LAST fail (earlier fails that get superseded by a later one don't
        # need their own separate recovery; only the final one matters).
        # 3-or-more consecutive FAILs at one tester is the more severe case,
        # checked regardless of whether it eventually recovered afterward.
        by_tester: dict[str, list[str]] = {}
        for e in ev:
            by_tester.setdefault(e.tester, []).append(e.result)
        status = "yes"
        for results in by_tester.values():
            run = max_run = 0
            for res in results:
                run = run + 1 if res == "FAIL" else 0
                max_run = max(max_run, run)
            if max_run > 2:
                status = "red"
                break
            last_fail = max((i for i, res in enumerate(results) if res == "FAIL"), default=-1)
            if last_fail != -1 and results[last_fail + 1:last_fail + 3] != ["PASS", "PASS"]:
                if status != "red":
                    status = "yellow"
        self.sop_status = status


def _is_ordered_subsequence(seq: list[str], target: list[str]) -> bool:
    it = iter(seq)
    return all(any(x == t for x in it) for t in target)


@dataclass
class ReportData:
    po: str
    build_code: str
    product: str
    cpn: str
    ordering_code: str
    period: str
    stations: list[str]
    all_devices: list[Device]        # every device found (for (Raw) Attempts)
    certified_devices: list[Device]  # final-pass devices (Device Detail, etc.)
    total_test_runs: int
    unique_devices: int
    final_pass: int
    final_fail: int
    success_rate: float
    failure_breakdown: list[tuple[str, int]]
    op_summary_final: dict           # operator -> {devices, pass, fail, stations, months}
    op_summary_all: dict             # operator -> {attempts, pass, fail, stations, months}
    daily_activity: dict             # date -> {operator: count}
    anomalies: list[Device]          # certified but did NOT complete all stages in order
    mc_retest: list[Device]          # heavy-retest devices
    third_stage_name: str            # whoever the raw data says does the final-cert stage
    # anomaly breakdown among certified devices
    n_completed_ordered: int = 0     # went Operator->OBA->third stage, in strict order
    n_oba_only: int = 0              # Operator + OBA only
    n_third_only: int = 0            # Operator + third stage only
    n_all3_unordered: int = 0        # all three present but not in strict order
    n_other_path: int = 0            # any other combination


def available_builds(attempts: list[Attempt]):
    """Summarise which (PO, Build Code, Product) combinations exist in the data,
    most common first — used to help the user when nothing matches."""
    counts = collections.Counter()
    for a in attempts:
        counts[(a.po or "—", a.build_code or "—", a.product or "—")] += 1
    return counts.most_common()


def build_report_data(attempts: list[Attempt], po: str, build_code: str,
                      mc_retest_min_attempts: int = 6) -> ReportData:
    po = (po or "").strip()
    build_code = (build_code or "").strip().upper()

    # ---- filter to the requested build (operator not checked yet) ------- #
    def po_build_match(a: Attempt) -> bool:
        if po and a.po != po:
            return False
        if build_code and a.build_code.upper() != build_code:
            return False
        return True

    candidates = [a for a in attempts if po_build_match(a)]

    # ---- discover this build's 3 stages from its OWN data --------------- #
    # "operator" and "OBA" are fixed role codes; the third stage is whoever
    # the raw data says it is, so a future rename needs no code change.
    third_name = _detect_third_stage(candidates) or _DEFAULT_THIRD_NAME
    operators = ["operator", "OBA", third_name]

    def canon(raw: str) -> Optional[str]:
        name = (raw or "").strip()
        low = name.lower()
        if low == "operator":
            return "operator"
        if low == "oba":
            return "OBA"
        return third_name if name else None

    sel = [a for a in candidates if canon(a.operator) is not None]

    # ---- de-duplicate identical records --------------------------------- #
    # Some raw files record every attempt twice; a real device never has two
    # distinct tests at the exact same timestamp by the same operator/station
    # with the same result, so exact-tuple duplicates are safe to drop.
    # (Records with no timestamp are left untouched.)
    seen: set = set()
    deduped: list[Attempt] = []
    for a in sel:
        if a.when is None:
            deduped.append(a)
            continue
        key = (a.device_id, a.when, canon(a.operator), a.station,
               a.passed, a.fail_reason)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(a)
    sel = deduped

    # ---- group into devices --------------------------------------------- #
    by_dev: dict[str, Device] = {}
    for a in sel:
        d = by_dev.get(a.device_id)
        if d is None:
            d = by_dev[a.device_id] = Device(imei=a.device_id)
        if not d.cpn and a.cpn:
            d.cpn = a.cpn
        if not d.ordering_code and a.ordering_code:
            d.ordering_code = a.ordering_code
        if not d.modem and a.modem:
            d.modem = a.modem
        d.events.append(TestEvent(
            tester=canon(a.operator),
            station=a.station,
            date=_fmt_date(a.when),
            time=_fmt_time(a.when),
            result="PASS" if a.passed else "FAIL",
            fail_reason=a.fail_reason,
            dev_id=a.dev_id,
            when=a.when,
        ))
    for d in by_dev.values():
        d.finalize(operators)

    all_devices = sorted(by_dev.values(), key=lambda d: (d.events[0].when if d.events else 0))
    certified = [d for d in all_devices if d.certified]

    # ---- headline metrics ----------------------------------------------- #
    total_runs = sum(d.total_attempts for d in all_devices)
    final_pass = len(certified)
    final_fail = len(all_devices) - final_pass
    success = final_pass / len(all_devices) if all_devices else 0.0

    # ---- failure reason breakdown (all failing attempts) ---------------- #
    fb = collections.Counter()
    for d in all_devices:
        for e in d.events:
            if e.result == "FAIL" and e.fail_reason:
                fb[e.fail_reason] += 1
    failure_breakdown = fb.most_common()

    # ---- operator summaries --------------------------------------------- #
    op_final = {op: {"devices": 0, "pass": 0, "fail": 0,
                     "stations": set(), "months": set()} for op in operators}
    for d in certified:
        op = d.final_operator
        if op in op_final:
            op_final[op]["devices"] += 1
            op_final[op]["pass"] += 1  # certified => pass
            for e in d.events:
                if e.tester == op:
                    op_final[op]["stations"].add(e.station)
                    if e.date:
                        op_final[op]["months"].add(e.date[:7])

    op_all = {op: {"attempts": 0, "pass": 0, "fail": 0,
                   "stations": set(), "months": set()} for op in operators}
    for d in all_devices:
        for e in d.events:
            if e.tester in op_all:
                s = op_all[e.tester]
                s["attempts"] += 1
                s["pass" if e.result == "PASS" else "fail"] += 1
                s["stations"].add(e.station)
                if e.date:
                    s["months"].add(e.date[:7])

    # ---- daily activity ------------------------------------------------- #
    # Per operator, count DISTINCT devices that operator passed, bucketed by the
    # date of that operator's last passing attempt for the device.
    # (Validated against the PO 1218 sample: operator 544, OBA 274, third-stage 273.)
    daily = collections.defaultdict(lambda: collections.Counter())
    for d in all_devices:
        last_pass_date = {}
        for e in d.events:
            if e.result == "PASS" and e.tester in operators and e.date:
                last_pass_date[e.tester] = e.date
        for op, dt in last_pass_date.items():
            daily[dt][op] += 1

    # ---- anomalies: certified but not strict operator->OBA->third stage - #
    anomalies = [d for d in certified if not d.completed_all_stages]

    # breakdown of the certified population by which stages each went through
    n_ord = n_oba = n_third = n_all3u = n_other = 0
    for d in certified:
        ops = {e.tester for e in d.events}
        has_o, has_b, has_t = "operator" in ops, "OBA" in ops, third_name in ops
        if has_o and has_b and has_t:
            if d.completed_all_stages:
                n_ord += 1
            else:
                n_all3u += 1
        elif has_o and has_b:
            n_oba += 1
        elif has_o and has_t:
            n_third += 1
        else:
            n_other += 1

    # ---- MC retest units ------------------------------------------------- #
    mc = [d for d in all_devices if d.total_attempts >= mc_retest_min_attempts]

    # ---- descriptive fields --------------------------------------------- #
    stations = sorted({e.station for d in all_devices for e in d.events if e.station})
    cpn = next((d.cpn for d in all_devices if d.cpn), "")
    ordering_code = next((d.ordering_code for d in all_devices if d.ordering_code), "")
    product = next((a.product for a in sel if a.product), "")
    months = sorted({m for d in all_devices for m in (d.months.split(", ") if d.months else [])})
    period = months[0] if len(months) == 1 else (f"{months[0]}..{months[-1]}" if months else "")

    return ReportData(
        po=po, build_code=build_code, product=product, cpn=cpn,
        ordering_code=ordering_code, period=period, stations=stations,
        all_devices=all_devices, certified_devices=certified,
        total_test_runs=total_runs, unique_devices=len(certified),
        final_pass=final_pass, final_fail=final_fail, success_rate=success,
        failure_breakdown=failure_breakdown,
        op_summary_final=op_final, op_summary_all=op_all,
        daily_activity=dict(daily), anomalies=anomalies, mc_retest=mc,
        third_stage_name=third_name,
        n_completed_ordered=n_ord, n_oba_only=n_oba, n_third_only=n_third,
        n_all3_unordered=n_all3u, n_other_path=n_other,
    )
