"""Compare extraction configs (model x output mode x reasoning effort) directly against OpenAI.

usage: python eval/compare.py <samples_dir> <expected.json> <out_dir> CONFIG [CONFIG ...] [--runs N]
CONFIG = model[:schema|json][:effort]   e.g. gpt-5.4-nano:json  gpt-5.4-mini:schema:low

Reads OPENAI_API_KEY from eval/.env (gitignored). Never prints it.
"""
import sys, os, re, json, time, pathlib, statistics, concurrent.futures as cf
import requests

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from candidate import SYSTEM, merge, load          # same prompt, cleanup and email loading
from score import score

def api_key():
    for line in (HERE / ".env").read_text(encoding="utf-8").splitlines():
        m = re.match(r'\s*OPENAI_API_KEY\s*=\s*"?([^"\s]+)"?', line)
        if m:
            return m.group(1)
    raise SystemExit("OPENAI_API_KEY not found in eval/.env")

KEY = api_key()
NULLABLE = lambda t: {"type": [t, "null"]}
SCHEMA = {
    "type": "json_schema",
    "json_schema": {"name": "vessels", "strict": True, "schema": {
        "type": "object", "additionalProperties": False, "required": ["vessels"],
        "properties": {"vessels": {"type": "array", "items": {
            "type": "object", "additionalProperties": False,
            "required": ["name", "dwt_k", "built", "open_location", "open_date"],
            "properties": {"name": {"type": "string"}, "dwt_k": NULLABLE("integer"), "built": NULLABLE("integer"),
                           "open_location": NULLABLE("string"), "open_date": NULLABLE("string")}}}}}},
}

def parse_config(c):
    parts = c.split(":")
    return {"name": c, "model": parts[0], "mode": parts[1] if len(parts) > 1 else "json",
            "effort": parts[2] if len(parts) > 2 else None}

def extract(cfg, p):
    subj, body = load(p)
    req = {"model": cfg["model"], "messages": [
        {"role": "system", "content": SYSTEM},
        {"role": "user", "content": f"Subject: {subj}\n\n{body}"}]}
    req["response_format"] = SCHEMA if cfg["mode"] == "schema" else {"type": "json_object"}
    if cfg["effort"]:
        req["reasoning_effort"] = cfg["effort"]
    t = time.perf_counter()
    out = {"file": p.name, "vessels": []}
    for attempt in range(3):
        try:
            r = requests.post("https://api.openai.com/v1/chat/completions", json=req,
                              headers={"Authorization": f"Bearer {KEY}"}, timeout=(10, 180))
            if r.status_code >= 400:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
            d = r.json()
            out["vessels"] = [{"name": v.get("name"), "dwt": f"{v['dwt_k']}K" if v.get("dwt_k") is not None else None,
                               "built": v.get("built"), "open": v.get("open_location"), "date": v.get("open_date")}
                              for v in merge(json.loads(d["choices"][0]["message"]["content"])["vessels"])]
            u = d.get("usage", {})
            out["in_tok"], out["out_tok"] = u.get("prompt_tokens", 0), u.get("completion_tokens", 0)
            out.pop("error", None)
            break
        except Exception as e:
            out["error"] = repr(e)[:300]
            time.sleep(2 ** attempt)
    out["seconds"] = round(time.perf_counter() - t, 2)
    return out

def main():
    args = sys.argv[1:]
    runs = 1
    if "--runs" in args:
        i = args.index("--runs"); runs = int(args[i + 1]); del args[i:i + 2]
    samples, expected_path, out_dir, configs = pathlib.Path(args[0]), args[1], pathlib.Path(args[2]), args[3:]
    expected = json.load(open(expected_path, encoding="utf-8"))
    files = sorted(samples.glob("*.eml"))
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"{'config':32} {'run':>3} {'correct':>8} {'found':>6} {'extra':>5} {'dwt':>4} {'blt':>4} {'open':>4} {'date':>4} "
          f"{'p50 s':>6} {'max s':>6} {'wall s':>6} {'in tok':>7} {'out tok':>7} err")
    for c in configs:
        cfg = parse_config(c)
        for run in range(1, runs + 1):
            t0 = time.perf_counter()
            with cf.ThreadPoolExecutor(8) as ex:
                res = list(ex.map(lambda p: extract(cfg, p), files))
            wall = time.perf_counter() - t0
            (out_dir / f"{c.replace(':', '_')}_run{run}.json").write_text(json.dumps(res, indent=1, ensure_ascii=False), encoding="utf-8")
            summary, _ = score(expected, res)
            n = lambda k: int(re.search(rf"{k} (\d+)", summary).group(1))
            found, tot = map(int, re.search(r"found (\d+)/(\d+)", summary).groups())
            secs = [r["seconds"] for r in res]
            errs = sum("error" in r for r in res)
            print(f"{c:32} {run:>3} {n('correct'):>4}/{tot:<3} {found:>6} {n('extras'):>5} {n('dwt'):>4} {n('built'):>4} {n('open'):>4} {n('date'):>4} "
                  f"{statistics.median(secs):>6.1f} {max(secs):>6.1f} {wall:>6.1f} {sum(r.get('in_tok',0) for r in res):>7} {sum(r.get('out_tok',0) for r in res):>7} {errs or ''}",
                  flush=True)
            for r in res:
                if "error" in r:
                    print("      error:", r["file"][:40], r["error"][:200])

if __name__ == "__main__":
    main()
