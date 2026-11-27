"""Check web/src/zones.ts gives the same zone as the desktop app's lookup_value() for a list of ports."""
import json, pathlib, subprocess, sys
ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "app"))
import logic
m = logic.load_csv_into_dict(logic.resource_path("WPIUpdated.csv"))
for k, v in logic.load_unlocode_dict().items():
    if k not in m: m[k] = v
    elif len(set(m[k])) > 1 and len(set(v)) == 1: m[k] = v
ports = ["SINGAPORE","BIN QASIM","QASIM BIN","CHITTAGONG","KUNSAN","GENOA","HOUSTON","ROTTERDAM","GEORGETOWN","GEORGETOWN(GY)",
 "PAULSBORO","PROVIDENCE","POINTE A PITRE","FREETOWN","FAZENDINHA","OTRANTO","DALIAN","CHAOZHOU","GUANGZHOU","HADIA","KANDLA",
 "NEW MANGALORE","GO-GIA","GO GIA","OKKE","CJK","PUERTO QUETZAL","RIZHAO","TIANJIN","DURBAN","CORTES","HO CHI MINH","HO CHI MIN",
 "KO SI CHANG","VIZAG","SIHANOUKVILLE","MOMBASA","HEREKE","LOME","DILISKELESI","DILISKELISI","DILISKELIESI","CALLAO","PORT OF SANTOS",
 "SANTOS","RECALADA","GIBRALTAR","PMO","USG","NOLA","QINGDAO","CAOFEIDIAN","JINGTANG","PORT KELANG","PORT KLANG","AQABA","IZMIR",
 "NEMRUT","ISKENDERUN","CONSTANTA","CONSTANZA","NOVOROSSIYSK","ODESSA","SAMARINDA","TABONEO","MUARA PANTAI","PARADIP","HALDIA",
 "KAKINADA","MUNDRA","TUTICORIN","DAMMAM","SOHAR","FUJAIRAH","JEBEL ALI","RICHARDS BAY","MAPUTO","TEMA","ABIDJAN","SAN NICOLAS",
 "PARANAGUA","ITAQUI","VANCOUVER","LONGVIEW","NEW ORLEANS","MOBILE","TAMPA","BALTIMORE","NEWCASTLE","PORT HEDLAND","GLADSTONE",
 "ZZZ NOWHERE","EAST CHINA","SOUTH CHINA","CAMBODIA","BRAZIL"]
py = {p: logic.lookup_value(p, m) for p in ports}
js_src = f"""
import {{ initZones, zoneFor }} from '{(ROOT / "web/src/zones.ts").as_uri()}';
import fs from 'node:fs';
initZones(JSON.parse(fs.readFileSync('{(ROOT / "web/public/zones.json").as_posix()}', 'utf8')));
const ports = {json.dumps(ports)};
console.log(JSON.stringify(Object.fromEntries(ports.map(p => [p, zoneFor(p)]))));
"""
tmp = pathlib.Path(sys.argv[1]) / "zone_parity.mts"
tmp.write_text(js_src.replace("import.meta.env", "({BASE_URL:'/'})"), encoding="utf-8")
out = subprocess.run(["node", str(tmp)], capture_output=True, text=True, encoding="utf-8")
if out.returncode:
    sys.exit(out.stderr[-2000:])
js = json.loads(out.stdout)
diff = [(p, py[p], js[p]) for p in ports if py[p] != js[p]]
for p, a, b in diff: print(f"DIFF {p!r}: desktop={a} web={b}")
print(f"{len(ports) - len(diff)}/{len(ports)} match; unknown in web: {[p for p in ports if js[p] == 'UNKNOWN']}")
