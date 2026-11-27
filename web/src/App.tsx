import { useEffect, useMemo, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import { initAuth, signIn, signOut } from "./auth";
import { listFolders, listMessages, type Folder, type Message } from "./graph";
import { extractVessels, looksRelevant, pool, type Vessel } from "./extract";
import { loadZones, setOverride, UNKNOWN, zoneCodes, zoneFor } from "./zones";
import { openDateKey } from "./dates";
import { downloadXlsx } from "./xlsx";
import { confirmCheckout, fetchStatus, LimitError, openBillingPortal, PLAN_INFO, type Status } from "./account";
import { Plans } from "./Plans";
import { Mark, Ping, ThemeToggle, Wake } from "./Brand";
import { currentTheme, setTheme } from "./theme";

type Row = Vessel & { key: string; received: string; from: string; subject: string; webLink: string };
type View = Row & { zone: string; dateKey: number | null };
type SortCol = "name" | "dwt_k" | "built" | "open_location" | "zone" | "dateKey" | "received" | "from";
type Sort = { col: SortCol; dir: 1 | -1 };
type Progress = { checked: number; relevant: number; done: number; failed: { subject: string; error: string }[] };

const PARALLEL = 6;
const day = (d: Date) => d.toISOString().slice(0, 10);
const fmtReceived = (iso: string) =>
  new Date(iso).toLocaleString(undefined, { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" });

export default function App() {
  const [phase, setPhase] = useState<"loading" | "out" | "in" | "error">("loading");
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [authError, setAuthError] = useState("");
  const [theme, setThemeState] = useState(currentTheme());
  const toggleTheme = () => {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    setThemeState(next);
  };

  useEffect(() => {
    initAuth()
      .then((a) => {
        setAccount(a);
        setPhase(a ? "in" : "out");
      })
      .catch((e) => {
        setAuthError(String(e?.errorCode || e?.message || e));
        setPhase("error");
      });
  }, []);

  return (
    <div className="shell">
      <header className="bar">
        <Mark size={36} />
        <span className="who">
          {account && (
            <>
              {account.username}
              <button className="link" onClick={() => signOut()}>Sign out</button>
            </>
          )}
          <ThemeToggle theme={theme} onToggle={toggleTheme} />
        </span>
      </header>
      {phase === "loading" && <p className="center muted">Loading…</p>}
      {phase === "error" && (
        <div className="center">
          <p>Sign-in didn't complete ({authError}).</p>
          <button className="primary" onClick={() => signIn()}>Try again</button>
        </div>
      )}
      {phase === "out" && <SignedOut />}
      {phase === "in" && <Workspace />}
    </div>
  );
}

function SignedOut() {
  return (
    <main className="welcome">
      <Mark size={88} />
      <h1 className="title">Welcome to Mail AI</h1>
      <p className="lede">Start getting vessel positions out of your inbox.</p>
      <button className="primary big" onClick={() => signIn()}>Sign in with Microsoft</button>
      <p className="small muted about"><a href="https://mailai.uk/">What is Mail AI?</a></p>
    </main>
  );
}

function Workspace() {
  const [folders, setFolders] = useState<Folder[] | null>(null);
  const [folderErr, setFolderErr] = useState("");
  const [folderId, setFolderId] = useState("");
  const [from, setFrom] = useState(day(new Date(Date.now() - 7 * 864e5)));
  const [to, setTo] = useState(day(new Date()));
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState<Progress | null>(null);
  const [rows, setRows] = useState<Row[]>([]);
  const [filter, setFilter] = useState("");
  const [sort, setSort] = useState<Sort>({ col: "received", dir: -1 });
  const [zoneFilter, setZoneFilter] = useState<Set<string>>(new Set());
  const [zonesReady, setZonesReady] = useState(false);
  const [zoneVersion, setZoneVersion] = useState(0); // bumps when the user overrides a port's zone
  const [editing, setEditing] = useState<string | null>(null); // open_location whose zone is being edited
  const [selected, setSelected] = useState<string | null>(null); // row key: the vessel you're on
  const [status, setStatus] = useState<Status | null>(null);
  const [limitHit, setLimitHit] = useState<LimitError["reason"] | null>(null);
  const [showPlans, setShowPlans] = useState(false);
  const [notice, setNotice] = useState<"paid" | null>(null);
  const stop = useRef(false);

  useEffect(() => {
    loadZones().finally(() => setZonesReady(true));
    const params = new URLSearchParams(window.location.search);
    const session = params.get("checkout");
    if (session) {
      history.replaceState(null, "", window.location.pathname); // don't re-confirm on reload
      confirmCheckout(session).then((s) => (setStatus(s), setNotice("paid"))).catch(() => fetchStatus().then(setStatus));
    } else {
      fetchStatus().then(setStatus).catch(() => {});
    }
  }, []);

  useEffect(() => {
    listFolders()
      .then((fs) => {
        setFolders(fs);
        setFolderId((fs.find((f) => f.depth === 0 && /^inbox$/i.test(f.name)) ?? fs[0])?.id ?? "");
      })
      .catch((e) => setFolderErr(String(e.message || e)));
  }, []);

  async function scan() {
    stop.current = false;
    setRunning(true);
    setLimitHit(null);
    setRows([]);
    setZoneFilter(new Set());
    const p: Progress = { checked: 0, relevant: 0, done: 0, failed: [] };
    setProgress({ ...p });
    const bump = () => setProgress({ ...p, failed: [...p.failed] });
    try {
      const until = new Date(to + "T00:00:00");
      until.setDate(until.getDate() + 1); // "to" date is inclusive
      for await (const page of listMessages(folderId, new Date(from + "T00:00:00"), until)) {
        if (stop.current) break;
        p.checked += page.length;
        const relevant = page.filter((m) => looksRelevant(m.subject, m.body));
        p.relevant += relevant.length;
        bump();
        await pool(relevant, PARALLEL, async (m: Message) => {
          if (stop.current) return;
          try {
            const vessels = await extractVessels(m.subject, m.body);
            const add = vessels.map((v, i) => ({
              ...v, key: `${m.id}:${i}`, received: m.received, from: m.from, subject: m.subject, webLink: m.webLink,
            }));
            setRows((r) => [...r, ...add]);
          } catch (e) {
            if (e instanceof LimitError) {
              // Trial over or daily cap reached: stop the whole scan, keep what we have.
              stop.current = true;
              setLimitHit(e.reason);
              setStatus(e.status);
              return;
            }
            p.failed.push({ subject: m.subject || "(no subject)", error: String((e as Error).message || e) });
          }
          p.done++;
          bump();
        });
      }
    } catch (e) {
      p.failed.push({ subject: "Reading mailbox", error: String((e as Error).message || e) });
      bump();
    } finally {
      setRunning(false);
      fetchStatus().then(setStatus).catch(() => {});
    }
  }

  const views: View[] = useMemo(
    () =>
      rows.map((r) => ({
        ...r,
        zone: zonesReady ? zoneFor(r.open_location, r.open_country) : "",
        dateKey: openDateKey(r.open_date, r.received),
      })),
    [rows, zonesReady, zoneVersion], // zoneVersion: recompute after an override
  );

  const zoneCounts = useMemo(() => {
    const c = new Map<string, number>();
    for (const v of views) c.set(v.zone, (c.get(v.zone) ?? 0) + 1);
    return [...c].sort((a, b) => b[1] - a[1]);
  }, [views]);

  const shown = useMemo(() => {
    const q = filter.trim().toUpperCase();
    const blank = (v: unknown) => v == null || v === "" || v === UNKNOWN;
    const { col, dir } = sort;
    return views
      .filter(
        (r) =>
          (zoneFilter.size === 0 || zoneFilter.has(r.zone)) &&
          (!q ||
            [r.name, r.open_location, r.open_country, r.zone, r.open_date, r.from, r.subject, r.dwt_k && `${r.dwt_k}K`, r.built]
              .some((v) => v != null && String(v).toUpperCase().includes(q))),
      )
      .sort((a, b) => {
        const x = a[col], y = b[col];
        if (blank(x) || blank(y)) return blank(x) === blank(y) ? 0 : blank(x) ? 1 : -1; // blanks always last
        const c = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y));
        return c * dir || b.received.localeCompare(a.received) || a.key.localeCompare(b.key);
      });
  }, [views, filter, zoneFilter, sort]);

  const toggleZone = (z: string) =>
    setZoneFilter((cur) => {
      const next = new Set(cur);
      if (next.has(z)) next.delete(z);
      else next.add(z);
      return next;
    });

  const th = (label: string, col: SortCol, num = false) => (
    <th className={num ? "num" : undefined} aria-sort={sort.col === col ? (sort.dir === 1 ? "ascending" : "descending") : "none"}>
      <button
        className="sort"
        onClick={() => setSort((s) => ({ col, dir: s.col === col ? (s.dir === 1 ? -1 : 1) : col === "received" ? -1 : 1 }))}
      >
        {label}
        <span className="arrow">{sort.col === col ? (sort.dir === 1 ? "▲" : "▼") : ""}</span>
      </button>
    </th>
  );

  if (folderErr) return <p className="center">Couldn't read your mail folders: {folderErr}</p>;

  return (
    <main className="work">
      <section className="controls">
        <label>
          Folder
          <select value={folderId} onChange={(e) => setFolderId(e.target.value)} disabled={!folders || running}>
            {!folders && <option>Loading folders…</option>}
            {folders?.map((f) => (
              <option key={f.id} value={f.id}>
                {"  ".repeat(f.depth) + f.name} ({f.total})
              </option>
            ))}
          </select>
        </label>
        <label>
          From
          <input type="date" value={from} max={to} onChange={(e) => setFrom(e.target.value)} disabled={running} />
        </label>
        <label>
          To
          <input type="date" value={to} min={from} onChange={(e) => setTo(e.target.value)} disabled={running} />
        </label>
        {running ? (
          <button onClick={() => (stop.current = true)}>Stop</button>
        ) : (
          <button
            className="primary" onClick={scan}
            disabled={!folderId || !from || !to || status?.plan === "expired"}
          >
            Extract positions
          </button>
        )}
        {status && <PlanLine s={status} onUpgrade={() => setShowPlans(true)} />}
      </section>

      {notice === "paid" && status && status.pro_source === "stripe" && (
        <p className="notice ok">
          Thank you for subscribing to {PLAN_INFO[status.plan as keyof typeof PLAN_INFO]?.name ?? status.plan}! A receipt is on its way to your inbox.
          <button className="link" onClick={() => setNotice(null)}>Dismiss</button>
        </p>
      )}
      {status?.plan === "expired" && <Paywall s={status} onUnlocked={setStatus} />}
      {showPlans && status && status.plan !== "expired" && (
        <Plans status={status} onPaid={(s) => (setStatus(s), setShowPlans(false), setNotice("paid"))} onClose={() => setShowPlans(false)} />
      )}
      {limitHit === "daily_cap" && status && (
        <p className="notice">
          You've reached today's limit of {status.daily_cap.toLocaleString()} position emails. Results so far are shown below.
          Your limit will reset at midnight UTC.
        </p>
      )}

      {progress && <ScanStatus p={progress} running={running} vessels={rows.length} />}

      {progress && !running && rows.length === 0 && progress.checked > 0 && (
        <section className="empty">
          <Wake />
          <p>No open positions found in {progress.relevant ? `${progress.relevant} position emails` : `${progress.checked} emails`}.</p>
          <p className="muted small">Try a wider date range or another folder.</p>
        </section>
      )}

      {rows.length > 0 && (
        <section className="results">
          <div className="results-bar">
            <input
              className="filter" placeholder="Search for a vessel, port, sender…" value={filter}
              onChange={(e) => setFilter(e.target.value)}
            />
            <span className="muted small">
              {shown.length === rows.length ? `${rows.length} positions` : `${shown.length} of ${rows.length} positions`}
            </span>
            <button onClick={() => downloadXlsx(shown, `Mail AI positions ${day(new Date())}.xlsx`)} disabled={!shown.length}>
              Export to Excel
            </button>
          </div>
          {zonesReady && zoneCounts.length > 1 && (
            <div className="chips" role="group" aria-label="Filter by zone">
              {zoneCounts.map(([z, n]) => (
                <button
                  key={z} className={zoneFilter.has(z) ? "chip on" : "chip"} aria-pressed={zoneFilter.has(z)}
                  onClick={() => toggleZone(z)}
                >
                  {z === UNKNOWN ? "No zone" : z} <span>{n}</span>
                </button>
              ))}
              {zoneFilter.size > 0 && <button className="link" onClick={() => setZoneFilter(new Set())}>Clear</button>}
            </div>
          )}
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  {th("Vessel", "name")}{th("DWT", "dwt_k", true)}{th("Built", "built", true)}{th("Open", "open_location")}
                  {th("Zone", "zone")}{th("Date", "dateKey")}{th("Received", "received")}{th("From", "from")}<th>Email</th>
                </tr>
              </thead>
              <tbody>
                {shown.map((r) => (
                  <tr
                    key={r.key} aria-selected={selected === r.key}
                    onClick={(e) => (e.target as HTMLElement).closest("a,button,select") || setSelected(r.key)}
                  >
                    <td className="strong">{r.name}</td>
                    <td className="num">{r.dwt_k != null ? `${r.dwt_k}K` : "–"}</td>
                    <td className="num">{r.built ?? "–"}</td>
                    <td>
                      {r.open_location ?? "–"}
                      {r.open_country && <span className="cc" title="Country">{r.open_country}</span>}
                    </td>
                    <td>
                      {editing === r.open_location && r.open_location ? (
                        <select
                          autoFocus value={r.zone === UNKNOWN ? "" : r.zone}
                          onChange={(e) => {
                            setOverride(r.open_location!, e.target.value || null);
                            setZoneVersion((v) => v + 1);
                            setEditing(null);
                          }}
                          onBlur={() => setEditing(null)}
                        >
                          <option value="">Automatic</option>
                          {zoneCodes().map((z) => <option key={z} value={z}>{z}</option>)}
                        </select>
                      ) : (
                        <button
                          className={r.zone === UNKNOWN ? "zone unknown" : "zone"}
                          title={r.open_location ? `Change zone for ${r.open_location}` : undefined}
                          disabled={!r.open_location}
                          onClick={() => setEditing(r.open_location)}
                        >
                          {r.zone === UNKNOWN ? "Set zone" : r.zone || "…"}
                        </button>
                      )}
                    </td>
                    <td className="nowrap">{r.open_date ?? "–"}</td>
                    <td className="muted nowrap">{fmtReceived(r.received)}</td>
                    <td className="muted clip">{r.from}</td>
                    <td className="clip"><a href={r.webLink} target="_blank" rel="noreferrer">{r.subject || "(no subject)"}</a></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  );
}

function ScanStatus({ p, running, vessels }: { p: Progress; running: boolean; vessels: number }) {
  const pct = p.relevant ? Math.round((100 * p.done) / p.relevant) : running ? 0 : 100;
  const plural = (n: number, one: string, many = `${one}s`) => `${n.toLocaleString()} ${n === 1 ? one : many}`;
  return (
    <section className="status" aria-live="polite">
      <div className="meter"><div style={{ width: `${pct}%` }} /></div>
      <p>
        {running ? (
          <>
            <Ping /> Extracting {p.done.toLocaleString()} of {plural(p.relevant, "position email")}
            <span className="muted"> · {plural(p.checked, "email")} checked</span>
          </>
        ) : (
          <>
            {plural(vessels, "position")} found from {plural(p.relevant, "email")}.
            <span className="muted"> {plural(p.checked, "email")} checked; the rest weren't position lists.</span>
          </>
        )}
        {p.failed.length > 0 && <span className="error"> {plural(p.failed.length, "email")} couldn't be read.</span>}
      </p>
      {p.failed.length > 0 && (
        <details>
          <summary>Show which</summary>
          <ul>{p.failed.map((f, i) => <li key={i}>{f.subject}: {f.error}</li>)}</ul>
        </details>
      )}
      {!running && p.checked === 0 && <p className="muted">No emails in that folder and date range.</p>}
    </section>
  );
}

function PlanLine({ s, onUpgrade }: { s: Status; onUpgrade: () => void }) {
  const usage = `${s.used_today.toLocaleString()} / ${s.daily_cap.toLocaleString()} today`;
  const paid = s.plan === "starter" || s.plan === "pro" || s.plan === "max";
  return (
    <div className="plan" title="Position emails extracted today (UTC). Other emails are skipped and don't count.">
      {paid && (
        <>
          <span className="badge pro">{PLAN_INFO[s.plan as keyof typeof PLAN_INFO].name}</span>
          <span className="muted">{usage}</span>
          {s.cancel_at_end && s.pro_until && <span className="muted">ends {new Date(s.pro_until).toLocaleDateString()}</span>}
          {s.can_manage ? (
            <button className="link" onClick={() => openBillingPortal()}>Manage</button>
          ) : (
            s.plan !== "max" && <button className="link" onClick={onUpgrade}>Upgrade</button>
          )}
        </>
      )}
      {s.plan === "trial" && (
        <>
          <span className="badge">Trial · {s.trial_days_left} {s.trial_days_left === 1 ? "day" : "days"} left</span>
          <span className="muted">{usage}</span>
          <button className="link" onClick={onUpgrade}>Upgrade</button>
        </>
      )}
      {s.plan === "expired" && <span className="badge off">Trial ended</span>}
    </div>
  );
}

function Paywall({ s, onUnlocked }: { s: Status; onUnlocked: (s: Status) => void }) {
  return (
    <section className="paywall">
      <h2>Your free trial has ended</h2>
      <p>Pick a plan to keep extracting positions.</p>
      <Plans status={s} onPaid={onUnlocked} />
    </section>
  );
}
