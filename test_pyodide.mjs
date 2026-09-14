// Runs test_sampling.py inside the same Pyodide build the browser uses,
// so the engine is verified against the exact pandas/openpyxl the site ships.
import { loadPyodide } from "pyodide";
import { readFileSync, existsSync } from "node:fs";

const py = await loadPyodide();
await py.loadPackage(["pandas", "micropip"]);
await py.pyimport("micropip").install("openpyxl==3.1.5");

const home = "/home/pyodide";
py.FS.writeFile(`${home}/qa_core.py`, readFileSync("public/qa_core.py"));
py.FS.writeFile(`${home}/bridge.py`, readFileSync("public/bridge.py"));
py.FS.writeFile(`${home}/test_sampling.py`, readFileSync("test_sampling.py"));
if (existsSync("sample_call_detail.xlsx")) {
  py.FS.writeFile(`${home}/sample_call_detail.xlsx`, readFileSync("sample_call_detail.xlsx"));
}

const code = py.runPython(`
import os, sys
os.chdir("${home}"); sys.path.insert(0, "${home}")
import test_sampling
test_sampling.main()
`);

// Bridge smoke check + selection fingerprint (compare with native CPython to prove seed parity).
const bridgeOk = py.runPython(`
import hashlib, json, os
import bridge, qa_core as core
if not os.path.exists("sample_call_detail.xlsx"):
    print("[skip] bridge check - sample_call_detail.xlsx not found"); ok = True
else:
    p = core.prepare_data(core.load_source("sample_call_detail.xlsx"), "sample_call_detail.xlsx")
    s = core.run_sampling(p, core.SamplingConfig(seed=20260820)).selected
    key = "\\n".join(s["Agent"].astype(str) + "|" + s["Sample Category"] + "|" + s["GenConnID"].astype(str))
    print("selection sha256 (seed 20260820):", hashlib.sha256(key.encode()).hexdigest())
    cfg = {"targets": [10, 5, 5], "seed": "20260820", "code_filter_enabled": False, "codes": "",
           "match_mode": "contains", "multi_code_logic": "any",
           "prioritize_voice_and_screen": True, "units": "auto"}
    a = json.loads(bridge.load("sample_call_detail.xlsx", "auto"))
    b = json.loads(bridge.sample(json.dumps(cfg), "2026-09-13T10:00:00"))
    bad = json.loads(bridge.sample(json.dumps({**cfg, "targets": [0, 0, 0]}), "2026-09-13T10:00:00"))
    ok = (a["ok"] and a["agents"] == 42 and b["ok"] and b["calls_selected"] == len(s)
          and b["filename"] == "QA_Random_Call_Sample_2026-09-13_100000.xlsx"
          and os.path.exists(b["path"]) and not bad["ok"])
    print("[PASS]" if ok else "[FAIL]", "bridge load/sample/validation", "" if ok else (a, bad))
ok
`);
process.exit(code || (bridgeOk ? 0 : 1));
