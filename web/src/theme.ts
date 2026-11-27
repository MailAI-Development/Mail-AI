/**
 * Chart (light) / Night watch (dark). Light by default; the user's choice is stored under the same key
 * the website uses, so mailai.uk and the app stay in step.
 */
export type Theme = "light" | "dark";
const KEY = "mailai-theme";

function stored(): Theme | null {
  try {
    const v = localStorage.getItem(KEY);
    return v === "light" || v === "dark" ? v : null;
  } catch {
    return null;
  }
}

export const currentTheme = (): Theme => stored() ?? "light";

export function applyTheme() {
  document.documentElement.setAttribute("data-theme", currentTheme());
}

export function setTheme(t: Theme) {
  try {
    localStorage.setItem(KEY, t);
  } catch {
    /* private mode: still switch for this page */
  }
  applyTheme();
}

/** Call once at start-up. */
export function initTheme() {
  applyTheme();
}
