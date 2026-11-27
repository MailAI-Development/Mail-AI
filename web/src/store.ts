/**
 * The vessel database: every email Mail AI has read, with the positions found in it, kept in this browser
 * (IndexedDB) and never on our server. One database per Microsoft account, so two people sharing a computer
 * don't see each other's positions. Emails with no positions are kept too, so they're never extracted twice.
 */
import type { Vessel } from "./extract";

export type StoredEmail = {
  id: string; // Graph message id
  received: string; // ISO
  from: string;
  subject: string;
  webLink: string;
  vessels: Vessel[];
};

const STORE = "emails";
const KEEP_KEY = "mailai.keepDays";
export const KEEP_OPTIONS = [7, 14, 30, 90];

export const keepDays = (): number => {
  const n = Number(localStorage.getItem(KEEP_KEY));
  return KEEP_OPTIONS.includes(n) ? n : 14;
};
export const setKeepDays = (n: number) => localStorage.setItem(KEEP_KEY, String(n));
export const keepSince = (days = keepDays()) => new Date(Date.now() - days * 864e5);

const req = <T>(r: IDBRequest<T>) =>
  new Promise<T>((resolve, reject) => {
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });

export class VesselStore {
  private constructor(private db: IDBDatabase) {}

  static open(accountId: string): Promise<VesselStore> {
    return new Promise((resolve, reject) => {
      const r = indexedDB.open(`mailai-${accountId}`, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: "id" }).createIndex("received", "received");
      r.onsuccess = () => resolve(new VesselStore(r.result));
      r.onerror = () => reject(r.error);
    });
  }

  private tx(mode: IDBTransactionMode) {
    return this.db.transaction(STORE, mode).objectStore(STORE);
  }

  all(): Promise<StoredEmail[]> {
    return req(this.tx("readonly").getAll() as IDBRequest<StoredEmail[]>);
  }

  /** Which of these message ids are already stored. */
  async known(ids: string[]): Promise<Set<string>> {
    const s = this.tx("readonly");
    const hits = await Promise.all(ids.map((id) => req(s.getKey(id))));
    return new Set(hits.filter((k): k is string => typeof k === "string"));
  }

  put(e: StoredEmail): Promise<unknown> {
    return req(this.tx("readwrite").put(e));
  }

  /** Delete emails received before `since`. Returns how many went. */
  async prune(since: Date): Promise<number> {
    const keys = await req(this.tx("readonly").index("received").getAllKeys(IDBKeyRange.upperBound(since.toISOString(), true)));
    if (!keys.length) return 0;
    const s = this.tx("readwrite");
    await Promise.all(keys.map((k) => req(s.delete(k))));
    return keys.length;
  }

  clear(): Promise<unknown> {
    return req(this.tx("readwrite").clear());
  }
}
