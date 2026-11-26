# Extraction eval

Measures how accurately vessel positions are extracted from real broker emails, so prompt or model
changes are judged by numbers, not by eye.

**Sample emails never go in this repo.** They live in `C:\dev\mailai-samples\` (`*.eml`) together with
the answer key `expected.json` (one entry per email: name, DWT in K, build year, accepted open ports,
accepted dates; fixed vessels, cargo orders and signature names are deliberately excluded).

## Scripts

| Script | What it does |
|---|---|
| `run_worker.py` | Scores the real `/v2/extract` endpoint (local `wrangler dev` or deployed). **Use this before deploying a change to `proxy/extract.js`.** |
| `compare.py` | Calls OpenAI directly to compare models / output modes / reasoning effort. Needs `OPENAI_API_KEY` in `eval/.env` (gitignored). |
| `candidate.py` | Python prototype of the prompt + trimming + cleanup that `compare.py` uses. Keep in sync with `proxy/extract.js`. |
| `baseline.py` | Runs the desktop app's v1.6 pipeline (`app/logic.py`) for comparison. |
| `score.py` | Scorer shared by all of the above. `--detail` lists every wrong field. |

## Typical loop

```sh
# terminal 1: run the Worker locally (reads proxy/.dev.vars: OPENAI_API_KEY, MAILAI_APP_SECRET)
cd proxy && npx wrangler@4 dev --port 8787

# terminal 2
.venv/Scripts/python -X utf8 eval/run_worker.py http://127.0.0.1:8787 C:/dev/mailai-samples C:/dev/mailai-samples/expected.json --runs 3
```

## Results (14 emails, 45 open vessels, Nov 2026)

| Pipeline | Fully correct | Time, all 14 | Cost / email |
|---|---|---|---|
| v1.6 desktop (nano, heavy line filtering, free-text output) | 23/45 (51%) | 32 s sequential | ~0.03¢ |
| v2 Worker (gpt-5.4-mini, light trimming, strict JSON schema) | 44–45/45 | ~3 s parallel | ~0.2¢ |

gpt-5.4-nano with the v2 prompt scored 40–42/45 and varied run to run; larger models (gpt-5.5, gpt-5.6-*)
were no more accurate and slower.
