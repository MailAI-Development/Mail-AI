"""Score an extraction run against expected.json.

usage: python score.py expected.json run.json [--detail]
run.json: [{"file": "<eml name>", "vessels": [{"MV"|"name", "Deadweight"|"dwt", "Build Year"|"built",
            "Vessel Open Location"|"open", "Vessel Open Date"|"date"}...]}, ...]
"""
import sys, json, re, difflib

MONTHS = {m[:3]: m[:3] for m in "JANUARY FEBRUARY MARCH APRIL MAY JUNE JULY AUGUST SEPTEMBER OCTOBER NOVEMBER DECEMBER".split()}

def g(v, *keys):
    for k in keys:
        if k in v:
            return v[k]
    return None

def norm_name(s):
    s = (s or "").upper()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"^\s*M\s*[./]?\s*V\s*[./]?\s+", "", s)
    s = re.sub(r"[\"'.,]", " ", s)
    return " ".join(s.split())

def norm_dwt(s):
    if s is None: return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*K", str(s).upper())
    if m: return round(float(m.group(1)))
    m = re.search(r"(\d[\d,]{3,})", str(s))
    if m: return round(int(m.group(1).replace(",", "")) / 1000)
    try: return round(float(s))
    except (TypeError, ValueError): return None

def norm_year(s):
    m = re.search(r"(19|20)\d{2}", str(s or ""))
    return int(m.group(0)) if m else None

def norm_loc(s):
    s = re.sub(r"[^A-Z ]", " ", (s or "").upper())
    return " ".join(s.split())

def norm_date(s):
    s = (s or "").upper()
    s = re.sub(r"\b(20\d{2})\b", " ", s)                 # drop years
    s = re.sub(r"(\d+)(ST|ND|RD|TH)\b", r"\1", s)        # ordinals
    s = re.sub(r"\b([A-Z]{3})[A-Z]*\b", lambda m: MONTHS.get(m.group(1), m.group(0)), s)
    s = re.sub(r"\b0(\d)\b", r"\1", s)
    s = re.sub(r"\s*[/\-–~]\s*", "-", s)
    s = re.sub(r"[.,]", " ", s)
    s = re.sub(r"\b(O-A|OA|ON-ABOUT|ABOUT|ABT|ARD|AROUND)\b", " ", s)
    s = " ".join(s.split()).strip("- ")
    s = re.sub(r"^(\d+)-([A-Z]{3})$", r"\1 \2", s)                       # 3-APR -> 3 APR
    s = re.sub(r"^(\d+) ([A-Z]{3})-(\d+) \2$", r"\1-\3 \2", s)          # 25 MAR-30 MAR -> 25-30 MAR
    return s

def name_sim(a, b):
    return difflib.SequenceMatcher(None, norm_name(a), norm_name(b)).ratio()

def loc_ok(pred, accepted):
    p = norm_loc(pred)
    return bool(p) and any(norm_loc(a) in p or p in norm_loc(a) for a in accepted)

def score(expected, run, detail=False):
    tot = dict(exp=0, pred=0, matched=0, dwt=0, built=0, open=0, date=0, perfect=0)
    lines = []
    for r in run:
        key = next((k for k in expected if not k.startswith("_") and r["file"].startswith(k)), None)
        if key is None:
            continue
        exp = expected[key]
        preds = [p for p in (r.get("vessels") or []) if norm_name(g(p, "MV", "name"))]
        tot["exp"] += len(exp); tot["pred"] += len(preds)
        used = set()
        lines.append(f"\n{r['file'][:70]}  ({len(preds)} predicted / {len(exp)} expected)")
        for e in exp:
            cands = [(name_sim(e["name"], g(p, "MV", "name")) + (0.2 if loc_ok(g(p, "Vessel Open Location", "open"), e["open"]) else 0), i)
                     for i, p in enumerate(preds) if i not in used and name_sim(e["name"], g(p, "MV", "name")) >= 0.8]
            if not cands:
                lines.append(f"   MISSING  {e['name']}")
                continue
            _, i = max(cands); used.add(i); p = preds[i]
            tot["matched"] += 1
            pd, pb = norm_dwt(g(p, "Deadweight", "dwt")), norm_year(g(p, "Build Year", "built"))
            ok = {
                "dwt": (pd is None) if e["dwt"] is None else (pd in e["dwt"]),
                "built": pb == e["built"],
                "open": loc_ok(g(p, "Vessel Open Location", "open"), e["open"]),
                "date": norm_date(g(p, "Vessel Open Date", "date")) in {norm_date(d) for d in e["date"]},
            }
            for k, v in ok.items(): tot[k] += v
            tot["perfect"] += all(ok.values())
            bad = [f"{k}: got {g(p, *{'dwt':('Deadweight','dwt'),'built':('Build Year','built'),'open':('Vessel Open Location','open'),'date':('Vessel Open Date','date')}[k])!r} want {e[k]!r}" for k, v in ok.items() if not v]
            lines.append(f"   {'ok  ' if not bad else 'FIX '}     {e['name']}" + ("   — " + "; ".join(bad) if bad else ""))
        for i, p in enumerate(preds):
            if i not in used:
                lines.append(f"   EXTRA    {g(p, 'MV', 'name')}")
    n = max(tot["exp"], 1)
    summary = (f"vessels found {tot['matched']}/{tot['exp']}  extras {tot['pred'] - tot['matched']}  "
               f"fully correct {tot['perfect']}/{tot['exp']} ({100*tot['perfect']/n:.0f}%)  |  "
               f"dwt {tot['dwt']}  built {tot['built']}  open {tot['open']}  date {tot['date']}  (of {tot['exp']})")
    return summary, lines

if __name__ == "__main__":
    expected = json.load(open(sys.argv[1], encoding="utf-8"))
    run = json.load(open(sys.argv[2], encoding="utf-8"))
    summary, lines = score(expected, run)
    if "--detail" in sys.argv:
        print("\n".join(lines))
    print(summary)
