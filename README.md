# QA Call Randomizer — Web

Browser version of the QA Call Randomizer v1.0.0 Windows desktop app. Upload a call detail
export, get a random per-agent QA sample, download the same 3-sheet Excel workbook
(**Selected Calls**, **Sampling Summary**, **Run Details**).

**Your file never leaves the browser.** The original Python engine (`qa_core.py`, unchanged)
runs locally via [Pyodide](https://pyodide.org) in a Web Worker. There is no backend.

All sampling rules — fixed column positions, duration buckets and gaps, Voice And Screen
priority, Post_Route_Data filter, duplicate prevention, no backfill — are documented in the
desktop README and behave identically. A **random seed from a desktop-generated workbook
reproduces the identical sample here** (verified: same selection fingerprint on native CPython
and in Pyodide for seed `20260820`).

## Files

| File | Purpose |
|---|---|
| `public/index.html` | The whole UI (replaces the PySide6 `app.py`) |
| `public/worker.js` | Loads Pyodide + pandas + openpyxl, runs the engine off the UI thread |
| `public/bridge.py` | JSON wrapper around `qa_core` (replaces the Qt workers) |
| `public/qa_core.py` | The desktop engine, byte-for-byte unchanged |
| `test_sampling.py` | Original 42-check suite |
| `test_pyodide.mjs` | Runs that suite + a bridge check inside the same Pyodide build the site uses |
| `make_sample_data.py` | Generates `sample_call_detail.xlsx` (12,500 rows / 42 agents) for testing |

## Develop

```bash
npm install
npm test          # 42 engine checks + bridge check, inside Pyodide (needs sample_call_detail.xlsx for E2E)
npm run dev       # http://localhost:8000
```

## Deploy

Static site, no build step. Vercel serves `public/`. Every push to `main` on GitHub redeploys.

## Differences from the desktop app

- **Download Excel Sample** saves through the browser instead of writing to `Documents\QA Call Randomizer\`; *Open Output Folder* is gone.
- First visit downloads the Python runtime (~15 MB, from `cdn.jsdelivr.net`, plus `openpyxl` from PyPI); later visits use the browser cache. Networks that block those hosts will show a load error.
- Run timestamps use the browser's local time.
