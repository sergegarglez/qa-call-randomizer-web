"""Test suite for the QA Call Randomizer core engine.

Run with:  python test_sampling.py
Covers Tests 1-10 from the specification plus duration normalisation.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import pandas as pd

import qa_core as core

PASS, FAIL = [], []


def check(name: str, condition: bool, detail: str = "") -> None:
    (PASS if condition else FAIL).append(name)
    mark = "PASS" if condition else "FAIL"
    print(f"[{mark}] {name}" + (f"  -> {detail}" if detail and not condition else ""))


# --------------------------------------------------------------------------
# helpers to build synthetic frames with the correct fixed column positions
# --------------------------------------------------------------------------
NCOLS = 45


def make_row(txn, agent, rec, when, duration, login, gen, post):
    row = [""] * NCOLS
    row[0], row[2], row[3] = txn, agent, rec
    row[9], row[11] = when, duration
    row[33], row[39], row[40] = login, gen, post
    return row


def frame(rows):
    cols = [f"Col{i}" for i in range(NCOLS)]
    cols[0], cols[2], cols[3] = "Transaction ID", "Agent Name", "Rec Type"
    cols[9], cols[11] = "Start Date Time", "Call Duration"
    cols[33], cols[39], cols[40] = "Login ID", "GenConnID", "Post_Route_Data"
    return pd.DataFrame(rows, columns=cols)


def build_agent(agent, n_short, n_mid, n_long, vas_ratio=1.0, code="ABC123", start=0):
    """n_* = calls in each duration bucket."""
    rows, i = [], start
    for count, (lo, hi) in ((n_short, (300, 1200)), (n_mid, (1260, 2100)), (n_long, (2160, 4000))):
        for k in range(count):
            secs = lo + (k * 7) % max(hi - lo, 1)
            rec = "Voice And Screen" if k < count * vas_ratio else "Voice Only"
            rows.append(make_row(
                f"TXN{i}", agent, rec, dt.datetime(2026, 8, 1, 9, 0) + dt.timedelta(minutes=i),
                core.format_hhmmss(secs), "U100", f"GC{i}", f"ROUTE={code}|TYPE=DIRECT"))
            i += 1
    return rows


def prep(rows, numeric_units="auto"):
    return core.prepare_data(frame(rows), "unit-test.xlsx", numeric_units)


def cfg(**kw):
    return core.SamplingConfig(**kw)


def bucket_counts(result, agent):
    sel = result.selected[result.selected["Agent"] == agent]
    return {b.label: int((sel["Sample Category"] == b.label).sum())
            for b in core.default_buckets()}


# --------------------------------------------------------------------------
# Duration normalisation
# --------------------------------------------------------------------------
def test_durations():
    cases = [
        ("05:32", 332), ("15:45", 945), ("22:10", 1330), ("41:20", 2480),
        ("00:41:20", 2480), ("01:02:03", 3723),
        (dt.time(0, 41, 20), 2480), (dt.timedelta(minutes=41, seconds=20), 2480),
        (2480 / 86400, 2480),            # Excel fractional day
        (round(0.0125, 6), 1080),        # 18:00 as fraction of a day
        (41.3333, 2480),                 # numeric minutes (auto)
        (2480.0, 2480),                  # numeric seconds (auto, >= 240)
        ("1h 05m 32s", 3932),
        ("", None), ("N/A", None), (None, None), (float("nan"), None), ("-5:00", None),
    ]
    ok = True
    for value, expected in cases:
        got = core.parse_duration(value)
        if expected is None:
            good = got is None
        else:
            good = got is not None and abs(got - expected) <= 1
        if not good:
            ok = False
            print(f"    duration mismatch: {value!r} -> {got} (expected {expected})")
    check("Duration normalisation (Excel time, fractions, HH:MM:SS, MM:SS, numeric)", ok)
    check("Explicit numeric units override",
          core.parse_duration(300, "seconds") == 300
          and core.parse_duration(300, "minutes") == 18000
          and core.parse_duration(0.5, "day_fraction") == 43200)
    check("Boundary formatting", core.format_hhmmss(2480) == "00:41:20")


# --------------------------------------------------------------------------
# Test 1 - agent with sufficient calls everywhere
# --------------------------------------------------------------------------
def test_1_full_sample():
    rows = build_agent("John Smith", 30, 20, 20)
    r = core.run_sampling(prep(rows), cfg(seed=42))
    counts = bucket_counts(r, "John Smith")
    check("Test 1 - full 20-call sample",
          r.calls_selected == 20 and list(counts.values()) == [10, 5, 5], str(counts))
    check("Test 1 - summary status Complete",
          r.summary.iloc[0]["Status"] == core.STATUS_COMPLETE)


# --------------------------------------------------------------------------
# Test 2 - insufficient short calls, no cross-bucket backfill
# --------------------------------------------------------------------------
def test_2_shortage():
    rows = build_agent("Ana Lopez", 6, 20, 20)
    r = core.run_sampling(prep(rows), cfg(seed=7))
    counts = bucket_counts(r, "Ana Lopez")
    check("Test 2 - shortage reported, no backfill from other buckets",
          counts["5\u201320 Minutes"] == 6 and counts["21\u201335 Minutes"] == 5
          and counts["36+ Minutes"] == 5 and r.calls_selected == 16, str(counts))
    check("Test 2 - warning raised", len(r.warnings) == 1 and "Ana Lopez" in r.warnings[0])
    check("Test 2 - status Partial", r.summary.iloc[0]["Status"] == core.STATUS_PARTIAL)


# --------------------------------------------------------------------------
# Test 3 - Voice And Screen prioritised
# --------------------------------------------------------------------------
def test_3_vas_priority():
    rows = build_agent("Luis Vega", 30, 20, 20, vas_ratio=1.0)
    rows += build_agent("Luis Vega", 20, 20, 20, vas_ratio=0.0, start=500)
    r = core.run_sampling(prep(rows), cfg(seed=99))
    sel = r.selected
    check("Test 3 - all selections are Voice And Screen when plentiful",
          (sel["Recording Priority"] == "Voice And Screen").all()
          and (sel["Selection Reason"] == core.REASON_VAS).all())


# --------------------------------------------------------------------------
# Test 4 - insufficient VAS -> random fallback
# --------------------------------------------------------------------------
def test_4_vas_fallback():
    rows = build_agent("Karla Ruiz", 7, 5, 5, vas_ratio=1.0)             # 7 VAS short calls
    rows += build_agent("Karla Ruiz", 15, 0, 0, vas_ratio=0.0, start=900)  # 15 other short calls
    r = core.run_sampling(prep(rows), cfg(seed=5))
    short = r.selected[r.selected["Sample Category"] == "5\u201320 Minutes"]
    vas = int((short["Recording Priority"] == "Voice And Screen").sum())
    other = int((short["Recording Priority"] == "Other Recording Type").sum())
    check("Test 4 - 7 Voice And Screen + 3 fallback",
          vas == 7 and other == 3, f"vas={vas} other={other}")
    check("Test 4 - fallback reason recorded",
          (short[short["Recording Priority"] == "Other Recording Type"]["Selection Reason"]
           == core.REASON_FALLBACK).all())
    r2 = core.run_sampling(prep(rows), cfg(seed=5, prioritize_voice_and_screen=False))
    short2 = r2.selected[r2.selected["Sample Category"] == "5\u201320 Minutes"]
    check("Test 4 - priority OFF treats all types equally",
          int((short2["Recording Priority"] == "Other Recording Type").sum()) > 3
          and (short2["Selection Reason"] == core.REASON_RANDOM).all())


# --------------------------------------------------------------------------
# Test 5 / 6 - Post_Route_Data filtering
# --------------------------------------------------------------------------
def test_5_6_code_filter():
    rows = build_agent("Diego Paz", 15, 10, 10, code="ABC123")
    rows += build_agent("Diego Paz", 15, 10, 10, code="ZZZ999", start=2000)
    p = prep(rows)

    r = core.run_sampling(p, cfg(seed=3, code_filter_enabled=True, codes=["ABC123"]))
    check("Test 5 - only matching records selected",
          r.selected["Post_Route_Data"].str.contains("ABC123").all()
          and r.calls_selected == 20)
    check("Test 5 - Code Match column populated",
          (r.selected["Code Match"] == "Matched").all())

    r_any = core.run_sampling(p, cfg(seed=3, code_filter_enabled=True,
                                     codes=["ABC123", "ZZZ999"], multi_code_logic="any"))
    check("Test 6 - Any Code widens the pool", r_any.calls_selected == 20)

    r_all = core.run_sampling(p, cfg(seed=3, code_filter_enabled=True,
                                     codes=["ABC123", "ZZZ999"], multi_code_logic="all"))
    check("Test 6 - All Codes finds nothing when no record has both",
          r_all.calls_selected == 0
          and r_all.summary.iloc[0]["Status"] == core.STATUS_NONE)

    r_exact = core.run_sampling(p, cfg(seed=3, code_filter_enabled=True,
                                       codes=["ROUTE=ABC123|TYPE=DIRECT"], match_mode="exact"))
    check("Test 6 - Exact match mode", r_exact.calls_selected == 20)
    r_exact_bad = core.run_sampling(p, cfg(seed=3, code_filter_enabled=True,
                                           codes=["ABC123"], match_mode="exact"))
    check("Test 6 - Exact rejects partial values", r_exact_bad.calls_selected == 0)

    r_case = core.run_sampling(p, cfg(seed=3, code_filter_enabled=True, codes=["abc123"]))
    check("Test 6 - matching is case-insensitive", r_case.calls_selected == 20)

    r_off = core.run_sampling(p, cfg(seed=3))
    check("Test 5 - filter OFF samples everything",
          (r_off.selected["Code Match"] == "Not Applicable").all())


# --------------------------------------------------------------------------
# Test 7 - duplicate GenConnID
# --------------------------------------------------------------------------
def test_7_duplicates():
    rows = build_agent("Sofia Marin", 30, 20, 20)
    dupes = []
    for row in rows[:10]:
        clone = list(row)
        clone[0] = clone[0] + "-COPY"
        dupes.append(clone)                     # same GenConnID, different Transaction
    r = core.run_sampling(prep(rows + dupes), cfg(seed=11))
    gen = r.selected["GenConnID"]
    check("Test 7 - no duplicate GenConnID in output", gen.duplicated().sum() == 0)
    check("Test 7 - no duplicate Transaction in output",
          r.selected["Transaction"].duplicated().sum() == 0)

    # boundary: a call must not be picked twice across buckets
    check("Test 7 - each call appears once overall", r.calls_selected == len(set(gen)))

    # blank GenConnID falls back to Transaction
    blank = [make_row("TXN-A", "Blank Gen", "Voice And Screen",
                      dt.datetime(2026, 8, 1, 9, 0), "00:10:00", "U1", "", "ROUTE=X")]
    p2 = prep(blank * 1 + [make_row("TXN-B", "Blank Gen", "Voice Only",
                                    dt.datetime(2026, 8, 1, 9, 5), "00:11:00", "U1", "", "ROUTE=X")])
    r2 = core.run_sampling(p2, cfg(seed=1))
    check("Test 7 - blank GenConnID falls back to Transaction", r2.calls_selected == 2)


# --------------------------------------------------------------------------
# Test 8 / 9 - invalid duration + blank agent
# --------------------------------------------------------------------------
def test_8_9_validation():
    rows = build_agent("Pablo Nieto", 12, 6, 6)
    rows.append(make_row("TXN-BAD", "Pablo Nieto", "Voice And Screen",
                         dt.datetime(2026, 8, 1, 9, 0), "N/A", "U1", "GC-BAD", "ROUTE=X"))
    rows.append(make_row("TXN-NOAGENT", "", "Voice And Screen",
                         dt.datetime(2026, 8, 1, 9, 0), "00:10:00", "U1", "GC-NA", "ROUTE=X"))
    p = prep(rows)
    v = p.validation
    check("Test 8 - invalid duration excluded and reported", v.invalid_duration == 1)
    check("Test 9 - blank agent excluded and reported", v.missing_agent == 1)
    r = core.run_sampling(p, cfg(seed=2))
    check("Test 8/9 - excluded rows never selected",
          "TXN-BAD" not in set(r.selected["Transaction"])
          and "TXN-NOAGENT" not in set(r.selected["Transaction"])
          and r.agents_processed == 1)


# --------------------------------------------------------------------------
# Test 10 - random seed reproducibility
# --------------------------------------------------------------------------
def test_10_seed():
    rows = build_agent("Elena Cruz", 40, 30, 30)
    p = prep(rows)
    a = core.run_sampling(p, cfg(seed=1234))
    b = core.run_sampling(p, cfg(seed=1234))
    c = core.run_sampling(p, cfg(seed=99))
    ids_a = sorted(a.selected["GenConnID"])
    ids_b = sorted(b.selected["GenConnID"])
    ids_c = sorted(c.selected["GenConnID"])
    check("Test 10 - same seed reproduces identical selection", ids_a == ids_b)
    check("Test 10 - different seed produces a different selection", ids_a != ids_c)
    d = core.run_sampling(p, cfg())
    check("Test 10 - blank seed auto-generates and records one",
          isinstance(d.seed_used, int) and d.seed_used > 0)
    check("Test 10 - selection is not simply the first rows",
          ids_a != sorted(x["GenConnID"] for x in
                          [dict(zip(("GenConnID",), (f"GC{i}",))) for i in range(20)]))


# --------------------------------------------------------------------------
# Duration boundary rules (the intentional gaps)
# --------------------------------------------------------------------------
def test_boundaries():
    edge = [
        ("TXN-299", 299), ("TXN-300", 300), ("TXN-1200", 1200), ("TXN-1230", 1230),
        ("TXN-1260", 1260), ("TXN-2100", 2100), ("TXN-2130", 2130), ("TXN-2160", 2160),
    ]
    rows = [make_row(t, "Edge Case", "Voice And Screen",
                     dt.datetime(2026, 8, 1, 9, 0), core.format_hhmmss(s), "U1", f"G{t}", "")
            for t, s in edge]
    r = core.run_sampling(prep(rows), cfg(seed=1))
    picked = set(r.selected["Transaction"])
    check("Boundaries - 20:30 and 35:30 excluded (intentional gaps)",
          "TXN-1230" not in picked and "TXN-2130" not in picked)
    check("Boundaries - 04:59 excluded, 05:00/20:00/21:00/35:00/36:00 included",
          "TXN-299" not in picked
          and {"TXN-300", "TXN-1200", "TXN-1260", "TXN-2100", "TXN-2160"} <= picked)


# --------------------------------------------------------------------------
# Agent grouping / case handling
# --------------------------------------------------------------------------
def test_grouping():
    rows = build_agent("John Smith", 12, 6, 6)
    rows += build_agent("  John  Smith ", 12, 6, 6, start=3000)
    rows += build_agent("JOHN SMITH", 12, 6, 6, start=4000)
    r = core.run_sampling(prep(rows), cfg(seed=8))
    check("Grouping - whitespace and case variants are one agent",
          r.agents_processed == 1 and r.calls_selected == 20)

    mixed = [make_row("T1", "Mia Soto", "voice and screen", dt.datetime(2026, 8, 1, 9, 0),
                      "00:10:00", "U1", "G1", ""),
             make_row("T2", "Mia Soto", " VOICE AND SCREEN ", dt.datetime(2026, 8, 1, 9, 0),
                      "00:11:00", "U1", "G2", "")]
    p = prep(mixed)
    check("Case handling - recording type comparison is case/space insensitive",
          p.validation.voice_and_screen == 2)


# --------------------------------------------------------------------------
# End-to-end on the generated sample file
# --------------------------------------------------------------------------
def test_end_to_end():
    sample = Path("sample_call_detail.xlsx")
    if not sample.exists():
        print("[skip] sample_call_detail.xlsx not found - run make_sample_data.py first")
        return
    import time
    t0 = time.time()
    df = core.load_source(sample)
    p = core.prepare_data(df, str(sample))
    load_s = time.time() - t0
    t1 = time.time()
    r = core.run_sampling(p, cfg(seed=20260820))
    sample_s = time.time() - t1
    check("E2E - agents detected", r.agents_processed == 42, str(r.agents_processed))
    check("E2E - no duplicate calls",
          r.selected["GenConnID"].duplicated().sum() == 0)
    check("E2E - every selection respects its bucket", all(
        (r.selected.loc[r.selected["Sample Category"] == b.label, "_duration_seconds"]
         .between(b.min_seconds, b.max_seconds if b.max_seconds else 10**9)).all()
        for b in core.default_buckets()))
    check("E2E - per-agent caps respected", all(
        r.summary[f"{b.label} Selected"].max() <= b.target for b in core.default_buckets()))
    check("E2E - source file untouched",
          sample.stat().st_size > 0 and core.load_source(sample).shape == df.shape)
    check("E2E - performance under 20s", load_s + sample_s < 20,
          f"load {load_s:.1f}s + sample {sample_s:.1f}s")
    out = core.write_output(r, "QA_Random_Call_Sample_TEST.xlsx")
    check("E2E - workbook written with 3 sheets",
          out.exists() and len(pd.ExcelFile(out).sheet_names) == 3)
    print(f"    -> {r.calls_selected}/{r.expected_total} calls, "
          f"{r.completion_pct}% complete, seed {r.seed_used}, "
          f"VAS {r.voice_and_screen_selected}/{r.calls_selected}, "
          f"load {load_s:.1f}s sample {sample_s:.1f}s, output {out.name}")
    if r.warnings:
        print(f"    -> {len(r.warnings)} shortage warning(s), e.g. {r.warnings[0]}")


def main() -> int:
    print("=" * 74)
    print(f"{core.APP_NAME} v{core.APP_VERSION} - core engine test suite")
    print("=" * 74)
    test_durations()
    test_1_full_sample()
    test_2_shortage()
    test_3_vas_priority()
    test_4_vas_fallback()
    test_5_6_code_filter()
    test_7_duplicates()
    test_8_9_validation()
    test_10_seed()
    test_boundaries()
    test_grouping()
    test_end_to_end()
    print("-" * 74)
    print(f"{len(PASS)} passed, {len(FAIL)} failed")
    if FAIL:
        for f in FAIL:
            print(f"  FAILED: {f}")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
