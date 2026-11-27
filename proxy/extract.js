/**
 * Vessel-position extraction (used by POST /v2/extract).
 *
 * Pipeline: trim boilerplate -> one OpenAI call with a strict JSON schema -> deterministic cleanup.
 * Tuned against the eval set in eval/ (see eval/README.md); change the prompt or trimming here,
 * then re-run the eval before deploying.
 */

export const SYSTEM_PROMPT = `You extract OPEN dry-bulk vessel positions from shipbroking emails. Return JSON only.

Output: {"vessels": [{"name", "dwt_k", "built", "open_location", "open_country", "open_date"}]}

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
- open_country: ISO 3166-1 alpha-2 code of the open port's country (e.g. "UA" for ODESSA, "MX" for MANZANILLO, MEXICO). Use the country written in the email; if none is written, give it only when the port name clearly identifies one country. null if unsure or if open_location is a region.
- open_date: day/range + 3-letter month, no year, e.g. "25 MAR", "26-27 MAR", "30 MAR-1 APR", "EARLY APR". Use the vessel's open/delivery date, not ETA at a later port.
- Never guess. Use null when a value is not in the email.`;

const nullable = (t) => ({ type: [t, "null"] });
export const RESPONSE_FORMAT = {
  type: "json_schema",
  json_schema: {
    name: "vessels",
    strict: true,
    schema: {
      type: "object",
      additionalProperties: false,
      required: ["vessels"],
      properties: {
        vessels: {
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["name", "dwt_k", "built", "open_location", "open_country", "open_date"],
            properties: {
              name: { type: "string" },
              dwt_k: nullable("integer"),
              built: nullable("integer"),
              open_location: nullable("string"),
              open_country: nullable("string"),
              open_date: nullable("string"),
            },
          },
        },
      },
    },
  },
};

// Lines that never carry name/DWT/built/open/date. Dropping them cuts ~40% of tokens.
const NOISE = new RegExp(
  "(bimco|iso[\\s-]*f?-?8217|marpol|sulphur content|liberty (of|to) us|good weather|beaufort|douglas sea|" +
    "mixing of bunkers|bunker (quality|spec|saving)|allowance of|knots?\\b.*\\b(lsfo|vlsfo|ifo|mgo|hsfo)|" +
    "\\b(idle|working|work)\\b.*\\bmt\\b|consum|ballast(ing)?/|hatch (size|cover)|tank top|grain capacity|bale|" +
    "^\\s*(mob|mobile|tel|fax|skype|wechat|whatsapp|e-?mail|teams|msn|qq|dir)\\b|mailto:|\\[cid:|https?://|" +
    "^\\s*no\\.?\\s*\\d\\s*[:)]|^\\s*h\\.?c\\.?\\s*\\d|^\\s*\\d-\\s*\\d|you don't often get email)",
  "i",
);
// ...unless the line mentions something we need.
const KEEP = /\b(m\.?v\.?|dwt|dwat|built|blt|build|delivery|open|o\/a|ard|eta|pass|fixed)\b|\d{2,3}\s*k\b/i;

export function trimBody(body) {
  return body
    .split(/\r?\n/)
    .map((l) => l.trimEnd())
    .filter((l) => l.trim() && (KEEP.test(l) || !NOISE.test(l)))
    .join("\n");
}

function cleanName(n) {
  n = (n || "").toUpperCase();
  n = n.replace(/\(.*?\)/g, " ");
  n = n.replace(/^\s*M\s*[./]?\s*V\s*[./]?\s+/, "");
  n = n.replace(/\b\d{2,3}(\.\d)?\s*K\b.*$/, ""); // trailing "57K/2012" etc.
  return n.replace(/"/g, " ").split(/\s+/).filter(Boolean).join(" ") || null;
}

function cleanDate(d) {
  if (!d) return null;
  d = d.toUpperCase().replace(/\b(ON\/ABOUT|ON\/ABT|O\/A|ABOUT|ABT|ARD|AROUND)\b/g, " ");
  d = d.replace(/(\d+)(ST|ND|RD|TH)\b/g, "$1").replace(/\s+TO\s+/g, "-");
  d = d.replace(/\s*[-–~]\s*/g, "-").replace(/\s+/g, " ").trim();
  return d.replace(/^(\d+) ([A-Z]{3})-(\d+) \2$/, "$1-$3 $2") || null; // "25 MAR-30 MAR" -> "25-30 MAR"
}

/** Normalise model output and collapse duplicates (same name + same/missing open port). */
export function cleanVessels(vessels) {
  const out = [];
  for (const raw of Array.isArray(vessels) ? vessels : []) {
    const v = {
      name: cleanName(raw.name),
      dwt_k: Number.isFinite(raw.dwt_k) ? raw.dwt_k : null,
      built: Number.isInteger(raw.built) && raw.built >= 1950 && raw.built <= 2100 ? raw.built : null,
      open_location: raw.open_location ? String(raw.open_location).toUpperCase().trim() || null : null,
      open_country: /^[A-Z]{2}$/.test(String(raw.open_country ?? "").trim().toUpperCase())
        ? String(raw.open_country).trim().toUpperCase()
        : null,
      open_date: cleanDate(raw.open_date),
    };
    if (!v.name) continue;
    if (v.dwt_k !== null && v.dwt_k >= 1000) v.dwt_k = Math.round(v.dwt_k / 1000); // tonnes, not K
    // Particulars header like "AMOYSAILING TBN-MV AMOY GALAXY OR MV AMOY APOLLO" -> the TBN entry
    const head = v.name.split(/\s*-\s*MV\b|\s+OR\s+/)[0].trim();
    if (head !== v.name && out.some((o) => o.name === head)) v.name = head;

    const twin = out.find(
      (o) => o.name === v.name && (!o.open_location || !v.open_location || o.open_location === v.open_location),
    );
    if (twin) {
      for (const k of Object.keys(v)) if (twin[k] == null && v[k] != null) twin[k] = v[k];
    } else {
      out.push(v);
    }
  }
  return out;
}
