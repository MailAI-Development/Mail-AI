/**
 * Mail AI extraction proxy (Cloudflare Worker)
 * --------------------------------------------
 * Holds the OpenAI key server-side so it never ships inside the desktop app.
 *
 *   app  ->  https://mailai.uk/v1/extract  (Netlify forwards /v1/*, see website/_redirects)
 *        ->  this Worker  ->  OpenAI chat completions
 *
 * Request:  POST /v1/extract   header X-App-Token: <APP_TOKEN baked into the app>
 *           body {"messages": [{role: "system"|"user", content: "..."}, ...]}   (max 2)
 * Response: {"content": "<model output>"}
 *
 * Secrets (Cloudflare dashboard → Worker → Settings → Variables, type "Secret",
 * or `npx wrangler secret put NAME`):
 *   OPENAI_API_KEY      the OpenAI key
 *   MAILAI_APP_SECRET   must equal APP_TOKEN in app/_build_secrets.py
 * Plain vars (optional, defaults in wrangler.toml):
 *   OPENAI_MODEL        default gpt-5.4-nano
 *   MAX_INPUT_CHARS     default 40000
 */

const OPENAI_URL = "https://api.openai.com/v1/chat/completions";
// Netlify gives up on a forwarded request after ~30s; fail cleanly before that.
const UPSTREAM_TIMEOUT_MS = 25000;

function json(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

// Constant-time comparison: hash both sides to equal-length digests, then XOR-compare.
async function safeEqual(a, b) {
  const enc = new TextEncoder();
  const [da, db] = await Promise.all([
    crypto.subtle.digest("SHA-256", enc.encode(a)),
    crypto.subtle.digest("SHA-256", enc.encode(b)),
  ]);
  const x = new Uint8Array(da), y = new Uint8Array(db);
  let diff = 0;
  for (let i = 0; i < x.length; i++) diff |= x[i] ^ y[i];
  return diff === 0;
}

async function handleExtract(request, env) {
  const token = request.headers.get("X-App-Token") || "";
  if (!env.MAILAI_APP_SECRET || !(await safeEqual(token, env.MAILAI_APP_SECRET))) {
    return json(401, { detail: "Unauthorized" });
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return json(400, { detail: "Invalid JSON" });
  }

  const messages = body && Array.isArray(body.messages) ? body.messages : [];
  if (messages.length === 0 || messages.length > 2) {
    return json(400, { detail: "Invalid request" });
  }
  for (const m of messages) {
    if (!m || typeof m !== "object" || !["system", "user"].includes(m.role) || typeof m.content !== "string") {
      return json(400, { detail: "Invalid request" });
    }
  }

  // Limit input size to control costs.
  const maxChars = parseInt(env.MAX_INPUT_CHARS || "40000", 10);
  const totalChars = messages.reduce((n, m) => n + m.content.length, 0);
  if (totalChars > maxChars) {
    return json(400, { detail: "Input too large" });
  }

  try {
    const resp = await fetch(OPENAI_URL, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: env.OPENAI_MODEL || "gpt-5.4-nano",
        messages: messages.map(({ role, content }) => ({ role, content })),
      }),
      signal: AbortSignal.timeout(UPSTREAM_TIMEOUT_MS),
    });
    if (!resp.ok) {
      console.log(`OpenAI error ${resp.status}: ${(await resp.text()).slice(0, 500)}`);
      return json(502, { detail: "Upstream error" });
    }
    const data = await resp.json();
    return json(200, { content: data.choices[0].message.content });
  } catch (e) {
    console.log(`OpenAI request failed: ${e}`);
    return json(502, { detail: "Upstream error" });
  }
}

export default {
  async fetch(request, env) {
    const { pathname } = new URL(request.url);

    if (request.method === "GET" && (pathname === "/health" || pathname === "/v1/health")) {
      return json(200, { ok: true });
    }
    if (pathname === "/v1/extract") {
      if (request.method !== "POST") return json(405, { detail: "Method not allowed" });
      return handleExtract(request, env);
    }
    return json(404, { detail: "Not found" });
  },
};
