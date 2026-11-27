"""Generate web/public/zones.json from the desktop app's zone data, so both apps zone ports identically.

Sources (exactly as app/main.py builds them): app/assets/WPIUpdated.csv, then UN/LOCODE seaports
mapped by country/state, plus the hand-written region aliases and country names in app/logic.py.

Run after changing any of those:   .venv/Scripts/python -X utf8 web/scripts/build_zones.py
"""
import json, pathlib, sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "app"))
import logic  # noqa: E402

mapping = logic.load_csv_into_dict(logic.resource_path("WPIUpdated.csv"))
for k, v in logic.load_unlocode_dict().items():           # same merge rule as app/main.py
    if k not in mapping:
        mapping[k] = v
    elif len(set(mapping[k])) > 1 and len(set(v)) == 1:
        mapping[k] = v

out = {
    "_generated_by": "web/scripts/build_zones.py — do not edit by hand",
    "zones": sorted(logic._ZONE_CODES),
    "countries": logic._COUNTRY_NAME_ZONE,
    "aliases": logic._REGION_ALIAS,
    "noise": logic._PORT_NOISE.pattern,
    # name -> zone, or name -> [zones] when a name is genuinely ambiguous
    "ports": {k: (sorted(set(v))[0] if len(set(v)) == 1 else sorted(set(v))) for k, v in sorted(mapping.items())},
}
dest = ROOT / "web" / "public" / "zones.json"
dest.parent.mkdir(parents=True, exist_ok=True)
dest.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
print(f"wrote {dest} ({len(out['ports'])} ports, {dest.stat().st_size // 1024} KB)")
