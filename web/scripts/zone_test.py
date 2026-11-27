"""Run web/src/zones.ts against zone_cases.json. Exit code 1 if any case fails.

usage: .venv/Scripts/python -X utf8 web/scripts/zone_test.py
Cases with expected null are printed for information only (zone not agreed yet).
"""
import json, os, pathlib, subprocess, sys, tempfile

ROOT = pathlib.Path(__file__).resolve().parents[2]
cases = json.loads((ROOT / "web/scripts/zone_cases.json").read_text(encoding="utf-8"))["cases"]
js = f"""
import {{ initZones, zoneFor }} from '{(ROOT / "web/src/zones.ts").as_uri()}';
import fs from 'node:fs';
initZones(JSON.parse(fs.readFileSync('{(ROOT / "web/public/zones.json").as_posix()}', 'utf8')));
const cases = {json.dumps(cases)};
console.log(JSON.stringify(cases.map(([p, c]) => zoneFor(p, c))));
"""
with tempfile.TemporaryDirectory() as tmp:
    f = pathlib.Path(tmp) / "zone_test.mts"
    f.write_text(js, encoding="utf-8")
    out = subprocess.run(["node", "--no-warnings", str(f)], capture_output=True, text=True, encoding="utf-8")
if out.returncode:
    sys.exit(out.stderr[-2000:])
got = json.loads(out.stdout)

failed = 0
for (port, cc, want), z in zip(cases, got):
    if want is None:
        print(f"  info     {port:26} {cc or '--':3} -> {z}")
    elif z != want:
        failed += 1
        print(f"  FAIL     {port:26} {cc or '--':3} -> {z:10} expected {want}")
checked = sum(1 for c in cases if c[2] is not None)
print(f"{checked - failed}/{checked} zone cases pass")
sys.exit(1 if failed else 0)
