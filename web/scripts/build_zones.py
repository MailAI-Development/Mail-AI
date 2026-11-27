"""Generate web/public/zones.json: port name (+ country) -> trading zone.

Sources, in priority order (highest first):
  1. web/scripts/zone_rules.json "ports"  - curated major ports in two-coast countries + English aliases
  2. app/assets/WPIUpdated.csv            - hand-maintained World Port Index list (name -> zone, no country)
  3. UN/LOCODE seaports (app/assets/csv)  - name + country (+ region); zone from zone_rules "subdivisions",
                                            then the country default (zone_rules "country_defaults" overrides
                                            app/logic.py; null = two coasts, no safe default)
Region aliases (CJK, SE ASIA, …) and country names come from app/logic.py, as in the desktop app.

Run after changing any of these:   .venv/Scripts/python -X utf8 web/scripts/build_zones.py
Then:                              .venv/Scripts/python -X utf8 web/scripts/zone_test.py
"""
import csv, json, os, pathlib, re, sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "app"))
import logic  # noqa: E402

rules = json.loads((ROOT / "web" / "scripts" / "zone_rules.json").read_text(encoding="utf-8"))
STATE_MAPS = {"US": logic._US_STATE_ZONE, "IN": logic._IN_STATE_ZONE, "CA": logic._CA_STATE_ZONE, "MX": logic._MX_STATE_ZONE}


def country_zone(cc: str, sub: str):
    if sub and sub in rules["subdivisions"].get(cc, {}):
        return rules["subdivisions"][cc][sub]
    if cc in STATE_MAPS:
        return STATE_MAPS[cc].get(sub)
    if cc in rules["country_defaults"]:
        return rules["country_defaults"][cc]
    return logic._COUNTRY_ZONE.get(cc)


def name_variants(name: str):
    """'DUMYAT (DAMIETTA)' -> DUMYAT (DAMIETTA), DUMYAT, DAMIETTA."""
    yield name
    m = re.match(r"^(.*?)\s*\((?:EX\s+)?([^)]*)\)\s*$", name)
    if m:
        for part in m.groups():
            part = part.strip()
            if len(part) >= 3:
                yield part


# UN/LOCODE seaports: name -> {country: zone}
un: dict[str, dict[str, str]] = defaultdict(dict)
country_zones: dict[str, set] = defaultdict(set)
for i in range(1, 4):
    with open(os.path.join(logic.resource_path("csv"), f"UNLOCODE CodeListPart{i}.csv"), encoding="latin-1") as f:
        for row in csv.reader(f):
            if len(row) < 7 or not row[6].strip().startswith("1"):
                continue
            cc, name, sub = row[1].strip().upper(), row[4].strip().upper(), row[5].strip().upper()
            if not cc or not name:
                continue
            zone = country_zone(cc, sub)
            if not zone:
                continue
            country_zones[cc].add(zone)
            for n in name_variants(name):
                un[n].setdefault(cc, zone)

# Curated ports win over UN/LOCODE for their country.
curated: dict[str, dict[str, str]] = defaultdict(dict)
for cc, ports in rules["ports"].items():
    for name, zone in ports.items():
        curated[name.upper()][cc] = zone
        un[name.upper()][cc] = zone
        country_zones[cc].add(zone)

wpi = {k: sorted(set(v)) for k, v in logic.load_csv_into_dict(logic.resource_path("WPIUpdated.csv")).items()}

out = {
    "_generated_by": "web/scripts/build_zones.py from zone_rules.json + WPI + UN/LOCODE. Do not edit by hand.",
    "zones": sorted(logic._ZONE_CODES),
    "countries": logic._COUNTRY_NAME_ZONE,
    "aliases": logic._REGION_ALIAS,
    "noise": logic._PORT_NOISE.pattern,
    "curated": {k: v for k, v in sorted(curated.items())},
    "wpi": {k: (v[0] if len(v) == 1 else v) for k, v in sorted(wpi.items())},
    "un": {k: v for k, v in sorted(un.items())},
    "country_zones": {k: sorted(v) for k, v in sorted(country_zones.items())},
}
dest = ROOT / "web" / "public" / "zones.json"
dest.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
amb = sum(1 for k, v in un.items() if k not in wpi and k not in curated and len(set(v.values())) > 1)
print(f"wrote {dest.name}: {len(curated)} curated, {len(wpi)} WPI, {len(un)} UN/LOCODE names "
      f"({amb} ambiguous without a country), {dest.stat().st_size // 1024} KB")
