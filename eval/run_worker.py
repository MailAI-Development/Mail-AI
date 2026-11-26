"""Score the deployed (or `wrangler dev`) /v2/extract endpoint against the answer key.

usage: python eval/run_worker.py <base_url> <samples_dir> <expected.json> [--runs N] [--save out.json]
e.g.   python eval/run_worker.py http://127.0.0.1:8787 C:/dev/mailai-samples C:/dev/mailai-samples/expected.json

Sends the raw subject + plain-text body; trimming, prompt and cleanup happen in the Worker (proxy/extract.js).
Token comes from MAILAI_APP_TOKEN or app/_build_secrets.py. Never printed.
"""
import sys, os, json, time, email, pathlib, statistics, concurrent.futures as cf
from email import policy
import requests

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from score import score

def token():
    if os.environ.get("MAILAI_APP_TOKEN"):
        return os.environ["MAILAI_APP_TOKEN"]
    sys.path.insert(0, str(HERE.parent / "app"))
    import _build_secrets
    return _build_secrets.APP_TOKEN

def main():
    args = sys.argv[1:]
    opts = {}
    for flag in ("--runs", "--save"):
        if flag in args:
            i = args.index(flag); opts[flag] = args[i + 1]; del args[i:i + 2]
    base, samples, expected_path = args[0].rstrip("/"), pathlib.Path(args[1]), args[2]
    expected = json.load(open(expected_path, encoding="utf-8"))
    tok = token()
    files = sorted(samples.glob("*.eml"))

    def one(p):
        m = email.message_from_bytes(p.read_bytes(), policy=policy.default)
        payload = {"subject": str(m["subject"] or ""), "body": m.get_body(("plain",)).get_content()}
        t = time.perf_counter()
        r = requests.post(f"{base}/v2/extract", json=payload, headers={"X-App-Token": tok}, timeout=(10, 60))
        out = {"file": p.name, "seconds": round(time.perf_counter() - t, 2), "status": r.status_code}
        if r.ok:
            out["vessels"] = [{"name": v["name"], "dwt": f"{v['dwt_k']}K" if v["dwt_k"] is not None else None,
                               "built": v["built"], "open": v["open_location"], "date": v["open_date"]}
                              for v in r.json()["vessels"]]
        else:
            out["vessels"], out["error"] = [], r.text[:200]
        return out

    for run in range(1, int(opts.get("--runs", 1)) + 1):
        t0 = time.perf_counter()
        with cf.ThreadPoolExecutor(8) as ex:
            res = list(ex.map(one, files))
        summary, lines = score(expected, res)
        secs = [r["seconds"] for r in res]
        errs = [r for r in res if "error" in r]
        print(f"run {run}: {summary}  |  p50 {statistics.median(secs):.1f}s  wall {time.perf_counter() - t0:.1f}s"
              + (f"  ERRORS {len(errs)}" if errs else ""))
        for r in errs:
            print("   ", r["status"], r["file"][:40], r["error"])
        for l in lines:
            if any(k in l for k in ("FIX", "MISSING", "EXTRA")):
                print(l)
        if "--save" in opts:
            pathlib.Path(opts["--save"]).write_text(json.dumps(res, indent=1, ensure_ascii=False), encoding="utf-8")

if __name__ == "__main__":
    main()
