"""Run the exact v1.6 extraction pipeline over sample .eml files; save results + timings."""
import sys, os, json, time, email, pathlib
from email import policy
APP = r"C:\dev\mail-ai\app"
sys.path.insert(0, APP); os.chdir(APP)
import logic
samples, out = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
csv_dict = logic.merge_custom_zones(logic.load_csv_into_dict(logic.resource_path("WPIUpdated.csv")))
results = []
for p in sorted(samples.glob("*.eml")):
    m = email.message_from_bytes(p.read_bytes(), policy=policy.default)
    subj = str(m["subject"] or "")
    body = m.get_body(("plain",)).get_content()
    r = {"file": p.name, "subject": subj, "body_chars": len(body)}
    r["relevant"] = logic.is_relevant_email(subj, body)
    pre = logic.get_first_n_lines(body) if r["relevant"] else None
    r["pre_chars"] = len(pre) if pre else 0
    if pre:
        t = time.perf_counter()
        try:
            v = logic.extract_details_from_email(pre, csv_dict)
            r["vessels_raw"] = v
            r["vessels"] = logic.filter_data(v)
        except Exception as e:
            r["error"] = repr(e)
        r["seconds"] = round(time.perf_counter() - t, 1)
    results.append(r)
    print(f"{r.get('seconds','-'):>6}s  raw={len(r.get('vessels_raw',[])):>3} kept={len(r.get('vessels',[])):>3}  pre={r['pre_chars']:>6}/{r['body_chars']:<6} {p.name[:60]}", flush=True)
out.write_text(json.dumps(results, indent=1, ensure_ascii=False), encoding="utf-8")
print("total seconds:", round(sum(r.get("seconds", 0) for r in results), 1))
