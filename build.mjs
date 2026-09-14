// Self-hosts the Python runtime: copies Pyodide from node_modules and every wheel the app needs
// into public/pyodide/, verifying each wheel's sha256. Visitors then never contact a third-party
// host; the downloads below happen only here, at build time.
import { createHash } from "node:crypto";
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";

const SRC = "node_modules/pyodide";
const OUT = "public/pyodide";
const ROOT_PACKAGES = ["pandas", "xlrd"];

// Not shipped by Pyodide: pure-Python wheels from PyPI, pinned by hash.
const PYPI = [
  { name: "openpyxl", version: "3.1.5", imports: ["openpyxl"], depends: ["et-xmlfile"],
    file_name: "openpyxl-3.1.5-py2.py3-none-any.whl",
    sha256: "5282c12b107bffeef825f4617dc029afaf41d0ea60823bbb665ef3079dc79de2",
    url: "https://files.pythonhosted.org/packages/c0/da/977ded879c29cbd04de313843e76868e6e13408a94ed6b987245dc7c8506/openpyxl-3.1.5-py2.py3-none-any.whl" },
  { name: "et-xmlfile", version: "2.0.0", imports: ["et_xmlfile"], depends: [],
    file_name: "et_xmlfile-2.0.0-py3-none-any.whl",
    sha256: "7a91720bc756843502c3b7504c77b8fe44217c85c537d85037f0f536151b2caa",
    url: "https://files.pythonhosted.org/packages/c1/8b/5fe2cc11fee489817272089c4203e679c63b570a5aaeb18d852ae3cbba6a/et_xmlfile-2.0.0-py3-none-any.whl" },
];

const sha256 = (buf) => createHash("sha256").update(buf).digest("hex");
const version = JSON.parse(readFileSync(`${SRC}/package.json`, "utf8")).version;
const CDN = `https://cdn.jsdelivr.net/pyodide/v${version}/full/`;
const lock = JSON.parse(readFileSync(`${SRC}/pyodide-lock.json`, "utf8"));

rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
for (const f of ["pyodide.mjs", "pyodide.asm.mjs", "pyodide.asm.wasm", "python_stdlib.zip"]) {
  copyFileSync(`${SRC}/${f}`, `${OUT}/${f}`);
}

for (const { url, ...entry } of PYPI) {
  lock.packages[entry.name] = { ...entry, install_dir: "site", package_type: "package", unvendored_tests: false };
}

const needed = new Set();
const stack = [...ROOT_PACKAGES, ...PYPI.map((p) => p.name)];
while (stack.length) {
  const name = stack.pop();
  if (needed.has(name)) continue;
  if (!lock.packages[name]) throw new Error(`Package "${name}" is not in pyodide-lock.json`);
  needed.add(name);
  stack.push(...lock.packages[name].depends);
}

let total = 0;
for (const name of needed) {
  const { file_name, sha256: expected } = lock.packages[name];
  const url = PYPI.find((p) => p.name === name)?.url ?? CDN + file_name;
  let bytes = existsSync(`${SRC}/${file_name}`) ? readFileSync(`${SRC}/${file_name}`) : null;
  if (!bytes || sha256(bytes) !== expected) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`Download failed (${res.status}): ${url}`);
    bytes = Buffer.from(await res.arrayBuffer());
  }
  if (sha256(bytes) !== expected) throw new Error(`sha256 mismatch for ${file_name}`);
  writeFileSync(`${OUT}/${file_name}`, bytes);
  total += bytes.length;
}

writeFileSync(`${OUT}/pyodide-lock.json`, JSON.stringify(lock));
console.log(`public/pyodide: Pyodide ${version} + ${needed.size} wheels (${(total / 1e6).toFixed(1)} MB) — ${[...needed].sort().join(", ")}`);
