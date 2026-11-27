/**
 * Dev only: a made-up mailbox for trying the app without real position emails. Shown as the folder
 * "Sample emails (dev)" when running `npx vite` locally; never part of the production build.
 * Two emails "arrive" 90 seconds and 3 minutes after the page loads, to try listening.
 * Clear saved positions to run it again from scratch.
 */
import type { Folder, Message } from "./graph";

export const DEMO_FOLDER = "demo-folder";
export const demoFolder: Folder = { id: DEMO_FOLDER, name: "Sample emails (dev)", depth: 0, total: 6 };

const loaded = Date.now();
const ago = (hours: number) => new Date(loaded - hours * 3600e3).toISOString();
const after = (seconds: number) => new Date(loaded + seconds * 1000).toISOString();

const mails: (Omit<Message, "webLink"> & { at: string })[] = [
  {
    id: "demo-1", at: ago(72), received: ago(72), from: "chartering@northsea-demo.test",
    subject: "OPEN TONNAGE LIST - SUPRAMAX / ULTRAMAX",
    body: `Good morning,

Please find our open tonnage below:

MV OCEAN PEARL 58K DWT / BLT 2014 - OPEN SINGAPORE 02-05 DEC
MV ATLANTIC DAWN 63K DWT / BLT 2017 - OPEN SANTOS 10 DEC
MV KESTREL BAY 57K DWT / BLT 2011 - OPEN ISKENDERUN 05/08 DEC
MV NORDIC TERN 61K DWT / BLT 2016 - OPEN QINGDAO PPT

Pls revert with suitable cargoes.

Best regards
North Sea Chartering (demo)`,
  },
  {
    id: "demo-2", at: ago(26), received: ago(26), from: "ops@harbourline-demo.test",
    subject: "Handysize positions",
    body: `Dear all,

Our handysize positions:

1) MV OCEAN PEARL 58K/2014 OPEN PORT KLANG 06 DEC
2) MV SEA-STAR 38K/2012 OPEN GDANSK 08 DEC
3) MV CORAL SPRINT 37K/2013 OPEN PORT ELLWOOD 12 DEC
4) MV GREY HERON 35K/2009 OPEN PORT ELLWOOD 14 DEC
5) MV LUNA GRACE 34K/2010 OPEN 20 DEC, DELY PORT TBA

Thanks and regards
Harbourline Ops (demo)`,
  },
  {
    id: "demo-3", at: ago(5), received: ago(5), from: "sam@example-demo.test",
    subject: "Lunch on Friday?",
    body: "Hi, are you free for lunch on Friday? There's a new place near the office. Sam",
  },
  {
    id: "demo-4", at: ago(2), received: ago(2), from: "chartering@northsea-demo.test",
    subject: "RE: OPEN TONNAGE LIST - updated",
    body: `Updated positions:

MV SEASTAR 38K/2012 OPEN ROSTOCK 15 DEC
KESTREL BAY OPEN MERSIN 09 DEC

Rest as per previous list.

North Sea Chartering (demo)`,
  },
  {
    id: "demo-5", at: after(90), received: after(90), from: "ops@harbourline-demo.test",
    subject: "Fresh positions - pls note",
    body: `Fresh tonnage:

MV NORDIC TERN 61K/2016 OPEN BUSAN 11 DEC
MV IRON WILLOW 82K/2019 OPEN PARADIP 13 DEC

Harbourline Ops (demo)`,
  },
  {
    id: "demo-6", at: after(180), received: after(180), from: "chartering@northsea-demo.test",
    subject: "OCEAN PEARL - open",
    body: "MV OCEAN PEARL 58K/2014 NOW OPEN SINGAPORE 14 DEC\n\nNorth Sea Chartering (demo)",
  },
];

/** Same shape as graph.listMessages: newest first, only emails that have "arrived" by now. */
export async function* demoMessages(since: Date, until: Date): AsyncGenerator<Message[]> {
  const now = Date.now();
  yield mails
    .filter((m) => new Date(m.at).getTime() <= now && m.received >= since.toISOString() && m.received < until.toISOString())
    .sort((a, b) => b.received.localeCompare(a.received))
    .map(({ at: _, ...m }) => ({ ...m, webLink: "#" }));
}
