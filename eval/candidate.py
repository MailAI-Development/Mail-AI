"""Python prototype of the v2 extraction (prompt, trimming, cleanup), run via the v1 proxy.

The production version lives in proxy/extract.js; compare.py imports SYSTEM/merge/load from here
to try other models. Keep the two in sync when changing the prompt.

usage: python eval/candidate.py <samples_dir> <out.json> [workers]
"""
import sys, os, re, json, time, email, pathlib, concurrent.futures as cf
from email import policy
import requests

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "app"))

SYSTEM = """You extract OPEN dry-bulk vessel positions from shipbroking emails. Return JSON only.

Output: {"vessels": [{"name", "dwt_k", "built", "open_location", "open_date"}]}

Rules:
- Include every vessel offered as open/available (tonnage lists, "open", "O/A", "ARD", "ETA", ex-shipyard deliveries, newbuildings/TBN).
- EXCLUDE: vessels marked FIXED/fixed/clean fixed/subs; cargo orders and tonnage requirements; names that only appear in signatures, PIC lists or contact details.
- One entry per open position. A vessel's particulars/description section is NOT a separate vessel: merge it into that vessel's entry.
- If the same vessel is genuinely offered with two different open positions (different ports in different sections), output both.
- If a vessel's only position is "PASS/PASSING X <date>", that IS its open position. But a passing point or ETA listed AFTER the vessel's open port is itinerary, not a second position.
- A newbuilding/TBN offered as "TBN - MV A or MV B" is ONE vessel: use the TBN name as written in the summary list.
- The subject line is part of the email and may hold the only position details.
- Vessel details are often split: a summary line near the top (name, DWT, open port, date) and a "vessel particulars"/description section further down (build year, exact DWT). Combine them.
- name: vessel name only, uppercase, no "MV", no notes in brackets (e.g. "XIN HAI TONG 35", not "XIN HAI TONG 35 (SCRUBBER FITTED)").
- dwt_k: summer deadweight in thousands, rounded to nearest integer (56,669 -> 57). null if not stated.
- built: 4-digit build/delivery year of the VESSEL. Never use a date from the open position or the email date. null if not stated anywhere.
- open_location: the port (or region code like CJK/ECI if no port), uppercase, no country, no "OPEN"/"AT" prefix. Correct obvious misspellings.
- open_date: day/range + 3-letter month, no year, e.g. "25 MAR", "26-27 MAR", "30 MAR-1 APR", "EARLY APR". Use the vessel's open/delivery date, not ETA at a later port.
- Never guess. Use null when a value is not in the email."""

def clean_name(n):
    n = (n or "").upper()
    n = re.sub(r"\(.*?\)", " ", n)
    n = re.sub(r"^\s*M\s*[./]?\s*V\s*[./]?\s+", "", n)
    n = re.sub(r"\b\d{2,3}(\.\d)?\s*K\b.*$", "", n)          # trailing '57K/2012' etc.
    return " ".join(n.replace('"', " ").split()) or None

def merge(vessels):
    """Collapse duplicates: same name + same open (or one side missing open) -> one entry."""
    out = []
    for v in vessels:
        v["name"] = clean_name(v.get("name"))
        if not v["name"]:
            continue
        if isinstance(v.get("dwt_k"), (int, float)) and v["dwt_k"] >= 1000:   # model gave tonnes, not K
            v["dwt_k"] = round(v["dwt_k"] / 1000)
        # Particulars header like "AMOYSAILING TBN-MV AMOY GALAXY OR MV AMOY APOLLO" -> same vessel as "AMOYSAILING TBN"
        head = re.split(r"\s*-\s*MV\b|\s+OR\s+", v["name"])[0].strip()
        if head != v["name"] and any(o["name"] == head for o in out):
            v["name"] = head
        twin = next((o for o in out if o["name"] == v["name"] and
                     (not o.get("open_location") or not v.get("open_location") or
                      o.get("open_location") == v.get("open_location"))), None)
        if twin:
            for k, val in v.items():
                if twin.get(k) in (None, "") and val not in (None, ""):
                    twin[k] = val
        else:
            out.append(v)
    return out

# Boilerplate that never carries name/DWT/built/open/date: dropped to cut tokens (cost).
NOISE = re.compile(r"(?i)(bimco|iso[\s-]*f?-?8217|marpol|sulphur content|liberty (of|to) us|good weather|beaufort|douglas sea|"
                   r"mixing of bunkers|bunker (quality|spec|saving)|allowance of|knots?\b.*\b(lsfo|vlsfo|ifo|mgo|hsfo)|"
                   r"\b(idle|working|work)\b.*\bmt\b|consum|ballast(ing)?/|hatch (size|cover)|tank top|grain capacity|bale|"
                   r"^\s*(mob|mobile|tel|fax|skype|wechat|whatsapp|e-?mail|teams|msn|qq|dir)\b|mailto:|\[cid:|https?://|"
                   r"^\s*no\.?\s*\d\s*[:)]|^\s*h\.?c\.?\s*\d|^\s*\d-\s*\d|you don't often get email)")
KEEP = re.compile(r"(?i)\b(m\.?v\.?|dwt|dwat|built|blt|build|delivery|open|o/a|ard|eta|pass|fixed)\b|\d{2,3}\s*k\b")

def trim(body):
    return "\n".join(l for l in body.splitlines() if KEEP.search(l) or not NOISE.search(l))

def load(p):
    m = email.message_from_bytes(p.read_bytes(), policy=policy.default)
    subj = str(m["subject"] or "")
    body = m.get_body(("plain",)).get_content()
    body = "\n".join(l.rstrip() for l in body.splitlines() if l.strip())
    return subj, trim(body) if os.environ.get("MAILAI_TRIM", "1") == "1" else body

def extract(p):
    import _build_secrets as S   # PROXY_URL, APP_TOKEN (never printed)
    subj, body = load(p)
    msgs = [{"role": "system", "content": SYSTEM},
            {"role": "user", "content": f"Subject: {subj}\n\n{body}"}]
    t = time.perf_counter()
    for attempt in range(3):
        try:
            r = requests.post(f"{S.PROXY_URL}/v1/extract", json={"messages": msgs},
                              headers={"X-App-Token": S.APP_TOKEN}, timeout=(10, 120))
            r.raise_for_status()
            txt = r.json()["content"].strip()
            txt = re.sub(r"^```(?:json)?|```$", "", txt).strip()
            vessels = merge(json.loads(txt)["vessels"])
            break
        except Exception as e:
            err = repr(e)
            vessels = None
    out = {"file": p.name, "seconds": round(time.perf_counter() - t, 1), "chars": len(subj) + len(body)}
    if vessels is None:
        out["error"] = err
        vessels = []
    out["vessels"] = [{"name": v.get("name"), "dwt": f"{v['dwt_k']}K" if v.get("dwt_k") is not None else None,
                       "built": v.get("built"), "open": v.get("open_location"), "date": v.get("open_date")}
                      for v in vessels]
    return out

if __name__ == "__main__":
    files = sorted(pathlib.Path(sys.argv[1]).glob("*.eml"))
    workers = int(sys.argv[3]) if len(sys.argv) > 3 else 8
    t0 = time.perf_counter()
    with cf.ThreadPoolExecutor(workers) as ex:
        results = list(ex.map(extract, files))
    wall = time.perf_counter() - t0
    for r in results:
        print(f"{r['seconds']:>6}s  {len(r['vessels']):>3} vessels  {r['chars']:>6} chars  {r['file'][:55]}  {r.get('error','')}")
    print(f"wall clock {wall:.1f}s with {workers} parallel")
    pathlib.Path(sys.argv[2]).write_text(json.dumps(results, indent=1, ensure_ascii=False), encoding="utf-8")
