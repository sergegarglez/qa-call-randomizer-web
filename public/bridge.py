"""
QA Call Randomizer - browser bridge.

Replaces the PySide6 layer (app.py) for the web build. Runs inside Pyodide in a
Web Worker and exposes qa_core as JSON-in / JSON-out calls. Every data decision
still lives in qa_core.py, which is shipped unchanged.
"""
from __future__ import annotations

import datetime as dt
import json
import traceback
from pathlib import Path

import qa_core as core

OUT_DIR = Path("/tmp/out")
_state: dict = {}


def _reply(fn, generic: str) -> str:
    try:
        return json.dumps({"ok": True, **fn()})
    except core.QaRandomizerError as exc:
        return json.dumps({"ok": False, "message": str(exc)})
    except Exception:  # pragma: no cover - defensive, mirrors app.py workers
        return json.dumps({"ok": False, "message": generic, "details": traceback.format_exc()})


def _prepare(units: str) -> dict:
    p = core.prepare_data(_state["df"], _state["path"], units)
    _state["prepared"], _state["units"] = p, units
    v = p.validation
    date_range = None
    if p.date_range and p.date_range[0] is not None:
        try:
            date_range = [f"{p.date_range[0]:%Y-%m-%d}", f"{p.date_range[1]:%Y-%m-%d}"]
        except Exception:
            date_range = None
    return {
        "name": Path(p.source_path).name,
        "agents": len(p.agents),
        "date_range": date_range,
        "headers": p.headers,
        "validation": v.as_rows(),
        "total_rows": v.total_rows,
        "usable_rows": v.usable_rows,
        "missing_agent": v.missing_agent,
        "invalid_duration": v.invalid_duration,
        "duplicate_ids": v.duplicate_ids,
    }


def load(path: str, units: str) -> str:
    def run():
        _state["df"] = core.load_source(path)
        _state["path"] = path
        return _prepare(units)
    return _reply(run, "Unable to read this file. Please upload a valid Excel or CSV file.")


def reprepare(units: str) -> str:
    return _reply(lambda: _prepare(units), "Unable to re-read the file with these duration units.")


def sample(config_json: str, local_now: str, progress=None) -> str:
    def run():
        c = json.loads(config_json)
        buckets = [core.Bucket(b.label, b.min_seconds, b.max_seconds, int(t))
                   for b, t in zip(core.default_buckets(), c["targets"])]
        seed_text = str(c.get("seed") or "").strip()
        seed = None
        if seed_text:
            try:
                seed = int(seed_text)
            except ValueError:  # same rule as the desktop app
                seed = abs(hash(seed_text)) % (10 ** 9)
        on = bool(c["code_filter_enabled"])
        config = core.SamplingConfig(
            buckets=buckets,
            prioritize_voice_and_screen=bool(c["prioritize_voice_and_screen"]),
            code_filter_enabled=on,
            codes=core.parse_codes(c.get("codes", "")) if on else [],
            match_mode="exact" if c["match_mode"] == "exact" else "contains",
            multi_code_logic="all" if c["multi_code_logic"] == "all" else "any",
            seed=seed,
            numeric_duration_units=c["units"],
        )
        if "prepared" not in _state:
            raise core.QaRandomizerError("Upload a call detail file first.")
        if config.calls_per_agent == 0:
            raise core.QaRandomizerError("Set at least one duration bucket to a value above zero.")
        if on and not config.codes:
            raise core.QaRandomizerError(
                "The Post_Route_Data filter is ON but no codes were entered.\n\n"
                "Enter at least one code, or switch the filter OFF.")
        if _state.get("units") != config.numeric_duration_units:
            _prepare(config.numeric_duration_units)

        # Pyodide's clock is UTC; stamp the run with the browser's local time instead.
        now = dt.datetime.fromisoformat(local_now)
        result = core.run_sampling(_state["prepared"], config, progress)
        result.run_timestamp = now
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        for old in OUT_DIR.glob("*.xlsx"):
            old.unlink()
        out = core.write_output(result, OUT_DIR / core.default_output_name(now))

        rows = [{
            "agent": str(r["Agent"]),
            "buckets": [[int(r[f"{b.label} Selected"]), int(r[f"{b.label} Target"])] for b in buckets],
            "total": [int(r["Total Selected"]), int(r["Total Target"])],
            "status": str(r["Status"]),
        } for _, r in result.summary.iterrows()]

        return {
            "path": str(out),
            "filename": out.name,
            "agents_processed": result.agents_processed,
            "calls_analyzed": result.calls_analyzed,
            "calls_selected": result.calls_selected,
            "expected_total": result.expected_total,
            "completion_pct": result.completion_pct,
            "voice_and_screen": result.voice_and_screen_selected,
            "other": result.other_selected,
            "seed": str(result.seed_used),
            "warnings": result.warnings,
            "priority": config.prioritize_voice_and_screen,
            "code_filter": config.code_filter_label(),
            "codes": config.codes,
            "match_mode": config.match_mode.title(),
            "multi_code_logic": "Any Code" if config.multi_code_logic == "any" else "All Codes",
            "rows": rows,
        }
    return _reply(run, "An unexpected error occurred while generating the sample.")
