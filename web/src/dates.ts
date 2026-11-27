const MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
const PART = { EARLY: 5, BEG: 3, BEGINNING: 3, MID: 15, MIDDLE: 15, END: 25, LATE: 25 } as Record<string, number>;

/**
 * Sort key (ms) for an open date like "25 MAR", "26-27 MAR", "30 MAR-1 APR", "EARLY APR" or "APR".
 * Uses the start of the window. The year is the one that puts the date nearest the email's
 * received date (positions are rarely more than a few months out). null if unparseable.
 */
export function openDateKey(openDate: string | null, receivedIso: string): number | null {
  if (!openDate) return null;
  const s = openDate.toUpperCase();
  const m = s.match(/\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*\b/);
  if (!m) return null;
  const month = MONTHS.indexOf(m[1]);
  const before = s.slice(0, m.index);
  const dayMatch = before.match(/(\d{1,2})(?!.*\d)/) ?? before.match(/^\D*(\d{1,2})/);
  let day = dayMatch ? Number(dayMatch[1]) : 1;
  const firstNum = s.match(/^\s*(\d{1,2})/);
  if (firstNum && /^\s*\d{1,2}\s*[-/]/.test(s)) day = Number(firstNum[1]); // "26-27 MAR" -> 26
  const word = s.match(/\b(EARLY|BEG|BEGINNING|MID|MIDDLE|END|LATE)\b/);
  if (!dayMatch && word) day = PART[word[1]];
  const received = new Date(receivedIso);
  let best: number | null = null;
  for (const y of [received.getFullYear() - 1, received.getFullYear(), received.getFullYear() + 1]) {
    const t = new Date(y, month, Math.min(Math.max(day, 1), 31)).getTime();
    if (best === null || Math.abs(t - received.getTime()) < Math.abs(best - received.getTime())) best = t;
  }
  return best;
}
