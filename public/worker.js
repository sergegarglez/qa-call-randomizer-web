// Runs the Python engine (qa_core.py via bridge.py) in Pyodide, off the UI thread.
// The uploaded file only ever exists in this worker's memory; nothing is sent to a server.
// Module worker: importScripts() to a CDN is blocked by some browsers/proxies, import is not.
import { loadPyodide } from "https://cdn.jsdelivr.net/pyodide/v314.0.6/full/pyodide.mjs";

const HOME = "/home/pyodide";

const ready = (async () => {
  const py = await loadPyodide();
  await py.loadPackage(["pandas", "micropip", "xlrd"]);
  await py.pyimport("micropip").install("openpyxl==3.1.5");
  for (const f of ["qa_core.py", "bridge.py"]) {
    const res = await fetch(f, { cache: "no-cache" });
    if (!res.ok) throw new Error(`Could not fetch ${f}: HTTP ${res.status}`);
    py.FS.writeFile(`${HOME}/${f}`, await res.text());
  }
  py.runPython(`import os, sys; sys.path.insert(0, "${HOME}"); os.makedirs("/tmp/in", exist_ok=True)`);
  return { py, bridge: py.pyimport("bridge") };
})();

ready.then(
  () => postMessage({ type: "engine", ok: true }),
  (err) => postMessage({ type: "engine", ok: false, details: String(err) }),
);

onmessage = async ({ data }) => {
  const { id } = data;
  let engine;
  try {
    engine = await ready;
  } catch (err) {
    return postMessage({ id, ok: false, message: "The Python engine could not be loaded. Check your internet connection and reload the page.", details: String(err) });
  }
  const { py, bridge } = engine;

  if (data.type === "load") {
    const path = `/tmp/in/${data.name.replace(/[\\/]/g, "_")}`;
    py.FS.writeFile(path, new Uint8Array(data.bytes));
    const reply = JSON.parse(bridge.load(path, data.units));
    py.FS.unlink(path); // the parsed frame stays in memory; drop the raw copy
    return postMessage({ id, ...reply });
  }

  if (data.type === "units") {
    return postMessage({ id, ...JSON.parse(bridge.reprepare(data.units)) });
  }

  if (data.type === "sample") {
    const progress = (message, pct) => postMessage({ type: "progress", message, pct });
    const reply = JSON.parse(bridge.sample(data.config, data.now, progress));
    if (!reply.ok) return postMessage({ id, ...reply });
    const bytes = py.FS.readFile(reply.path);
    return postMessage({ id, ...reply, bytes: bytes.buffer }, [bytes.buffer]);
  }
};
