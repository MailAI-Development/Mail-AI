import lockup from "./brand/logos/mailai-lockup.svg";
import lockupReversed from "./brand/logos/mailai-lockup-reversed.svg";
import mark from "./brand/logos/mailai-mark.svg";
import markReversed from "./brand/logos/mailai-mark-reversed.svg";

/** The mark alone (the name is nearby). Reversed file on Night watch. */
export function Mark({ size = 64 }: { size?: number }) {
  return (
    <span className="lockup" role="img" aria-label="Mail AI">
      <img className="for-light" src={mark} alt="" width={size} height={size} />
      <img className="for-dark" src={markReversed} alt="" width={size} height={size} />
    </span>
  );
}

/** Mark + wordmark; the reversed file on Night watch (dark). */
export function Lockup({ height = 28 }: { height?: number }) {
  const width = Math.round((height * 196) / 64);
  return (
    <span className="lockup" role="img" aria-label="Mail AI">
      <img className="for-light" src={lockup} alt="" width={width} height={height} />
      <img className="for-dark" src={lockupReversed} alt="" width={width} height={height} />
    </span>
  );
}

/** The wake: the envelope's V-crease repeated. Empty states and dividers only, never behind text. */
export function Wake({ count = 5 }: { count?: number }) {
  const pitch = 24;
  return (
    <svg className="wake" width={count * pitch} height="14" viewBox={`0 0 ${count * pitch} 14`} aria-hidden="true">
      {Array.from({ length: count }, (_, i) => (
        <polyline key={i} points={`${i * pitch + 2},2 ${i * pitch + 12},11 ${i * pitch + 22},2`} />
      ))}
    </svg>
  );
}

/** Light/dark switch. Line icons, 1.75 stroke, 24px grid, currentColor (design system iconography). */
export function ThemeToggle({ theme, onToggle }: { theme: "light" | "dark"; onToggle: () => void }) {
  const toDark = theme === "light";
  return (
    <button className="icon" onClick={onToggle} aria-label={toDark ? "Switch to dark theme" : "Switch to light theme"} title={toDark ? "Dark theme" : "Light theme"}>
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
        {toDark ? (
          <path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5Z" />
        ) : (
          <>
            <circle cx="12" cy="12" r="4" />
            <path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4" />
          </>
        )}
      </svg>
    </button>
  );
}

/** The position ping: something live or new. */
export const Ping = () => <span className="ping" aria-hidden="true" />;
