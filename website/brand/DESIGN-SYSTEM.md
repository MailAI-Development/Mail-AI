# Mail AI — design system

Mail AI reads shipbrokers' position emails and turns them into a clean list of open vessels. The whole brand comes from that one move: **a ship sailing out of an envelope**. Every rule below should make the product look calm, exact and built for people who read fifty position lists before lunch.

Three ideas hold it together:

- **Two foundations, one signal.** Warm chart-paper (`chart`) and deep hull navy (`hull`) do almost all the work. Buoy orange (`signal`) is the one colour that pops, and it is rare.
- **One typeface.** Switzer for everything: logo, headings, tables, buttons. Hierarchy comes from size and weight, never from a second family.
- **The sea is a hint, not a theme.** No waves, anchors, ropes, compass roses or stock ships. The maritime feel comes from the colour names, the waterline hairline, the wake chevron and the position ping in the logo, and that's all.

## Voice

Mail AI talks like a good ops colleague: short and literal, and it speaks the trade's language without explaining it.

- Do: state results plainly: "38 positions found from 14 emails."
- Don't: sea puns ("smooth sailing!", "all hands on deck"), exclamation marks, emoji, or AI hype words (*magic*, *supercharge*, *revolutionise*, *seamless*).
- Don't: call it "the AI". It's Mail AI, or just "extraction".
- Sentence case everywhere, including buttons: "Extract positions", "Export to Excel".
- No em dashes anywhere, en at most

## The logo

**Mark.** A bulk carrier whose hull is an envelope. The V-crease across the hull is the envelope's flap, the superstructure sits at the stern the way it does on a real dry-bulk ship, and the orange dot ahead of the bridge is the **position ping**: a vessel located, an unread message, the one bit of signal in the logo.

| File | Use |
| --- | --- |
| `mailai-lockup.svg` | Default. Mark plus wordmark on `chart` or white. |
| `mailai-lockup-reversed.svg` | On `hull` navy or a dark photo. |
| `mailai-mark.svg` | Where the name already appears nearby: favicon, app header, avatar. |
| `mailai-mark-reversed.svg` | Mark on navy. |
| `mailai-mark-mono.svg` | One ink (`currentColor`): embossing, fax-grey documents, single-colour print. |
| `mailai-wordmark.svg` | Wordmark only (`currentColor`), for tight horizontal spaces. |
| `mailai-app-icon.svg` | Windows installer, taskbar, Gumroad thumbnail. Navy tile, `radius-lg`. |

Rules:

- Clear space around the lockup = the height of the position ping × 2 on every side.
- Minimum size: mark 16px; lockup 96px wide. Below 20px, drop the ping's colour and use the mono mark.
- The ping is always `signal` (or `currentColor` in mono). Never recolour the hull orange, never add a second dot.
- Never outline, add shadows to, rotate or animate the mark. The ship doesn't bob.
- The wordmark is outlined Switzer Semibold (600) at −2% tracking. Don't retype it in another weight.

## Colour

Light theme is **Chart**, dark is **Night watch**. They swap the two foundations: the page goes from paper to navy and the ink goes from navy to paper, so the brand reads the same either way.

| Token | Light | Role |
| --- | --- | --- |
| `chart` | `#F4F1EA` | Page. Foundation 1. |
| `chart-raised` / `chart-sunk` | `#FBFAF6` / `#E9E4D8` | Panels on top / wells below |
| `hull` | `#0F1E2E` | Text and logo. Foundation 2. |
| `hull-muted` | `#4A5868` | Secondary text |
| `line` / `line-strong` | `#D6D0C2` / `#7C8289` | Hairlines / control borders |
| `signal` | `#E4572E` | The accent fill |
| `signal-ink` | `#B5401C` | Orange as text |
| `state-ok` / `state-error` | `#2A7352` / `#B3261E` | States only |

Proportion: roughly **70% chart, 25% hull, under 5% signal**. If a screen looks orange, it's wrong.

`signal` is for exactly these: the one primary button in a view, the focus ring, the position ping, a "new" or "live" marker, and the active-row indicator in a positions table. Orange text uses `signal-ink`, never `signal` (3.3:1 on paper is too faint). Text on an orange fill is always `on-signal` (navy), never white.

`state-ok` and `state-error` only mean state. They never decorate, and they always come with a word or an icon.

## Type

Switzer by the Indian Type Foundry, three static files in `fonts/`: `Switzer-400.woff2`, `Switzer-500.woff2`, `Switzer-600.woff2` (latin subset). It's free for personal and commercial use under the ITF Free Font License, from Fontshare (fontshare.com), not Google Fonts; get the full files there for the desktop app. A neutral Swiss-style grotesk: it lets the data and the mark do the talking. No monospace, not even for IMO numbers or coordinates. Turn on tabular figures instead (`font-variant-numeric: tabular-nums`) anywhere numbers line up in columns: DWT, dates, prices.

| Style | Size / line | Weight | Use |
| --- | --- | --- | --- |
| `display` | 56 / 60 | 600, −2.5% | Website hero only |
| `title` | 32 / 38 | 600, −2% | Page titles |
| `heading` | 22 / 28 | 600 | Section headings, dialog titles |
| `subheading` | 17 / 24 | 600 | Card titles, vessel names in detail view |
| `body` | 15 / 22 | 400 | Default text, table cells |
| `body-strong` | 15 / 22 | 500 | Vessel names in tables, emphasis |
| `small` | 13 / 18 | 400 | Captions, timestamps, email sender |
| `label` | 12 / 16 | 600, +6%, caps | Column headers, field labels |

Weights in use: 400, 500, 600. Never 700 or above in the product; bold-bold looks like slop.

## Layout and shape

- Spacing runs on a 4px base: `space-1` 4 · `space-2` 8 · `space-3` 12 · `space-4` 16 · `space-6` 24 · `space-8` 32 · `space-12` 48.
- Corners are crisp, not bubbly: `radius-md` (6px) is the default for buttons, inputs and cards; `radius-sm` (3px) for badges; `radius-lg` (12px) for dialogs.
- **Borders, not shadows.** Separate surfaces with a 1px `line` hairline and a step between `chart` and `chart-raised`. Shadows only on things that float (menus, dialogs), and then one soft one.
- Tables are the product. Rows 40px, `label` headers on `chart-sunk`, `line` rules between rows, no zebra stripes. The selected row gets a 3px `signal` bar on its leading edge (a single exception to "no coloured left borders", because it marks *the vessel you're on*).

## Maritime hints (keep them subtle)

These are the only nautical details allowed. Each has one job.

1. **The waterline.** A single `line` hairline under the app header and above the footer on the website. It's just a divider, but it sits where a waterline would.
2. **The wake.** The envelope's V-crease, repeated as a row of 3–6 small chevrons in `line` or `hull-muted`. Use it as a section divider on the website or an empty-state ornament. Never as a background pattern or behind text.
3. **The position ping.** A 6–8px `signal` dot marks something live or new: a freshly extracted position, an unread email, the sync indicator. It's the logo's dot doing its job in the UI.
4. **The names.** `chart`, `hull`, `signal`. Use them in code and in conversation.

Anything else (waves, anchors, ship's wheels, rope borders, compass roses, nautical-flag alphabets, ocean gradients) is out.

## Iconography

Line icons at 1.75px stroke on a 24px grid, round caps, drawn in `currentColor`, coloured `hull` or `hull-muted`. Lucide matches this style if you need a set, but that's a suggestion, not something in this system yet. Icons are never orange except the position ping.

## Using this system

Colours are semantic tokens with a light (Chart) and a dark (Night watch) theme. Light is the first theme and the fallback. Load Switzer from `fonts/` or Fontshare's CSS (`https://api.fontshare.com/v2/css?f[]=switzer@400,500,600`). The mono mark and wordmark are drawn in `currentColor`, so they recolour with CSS `color`; the two-colour marks have their colours baked in.

No components are defined yet: buttons, inputs, the positions table and badges are described above in words and tokens, ready to build.