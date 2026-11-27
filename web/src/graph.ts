import { graphToken } from "./auth";

const GRAPH = "https://graph.microsoft.com/v1.0";

async function get<T>(pathOrUrl: string): Promise<T> {
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : GRAPH + pathOrUrl;
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${await graphToken()}`,
        Prefer: 'outlook.body-content-type="text"', // plain-text bodies, like the desktop app
      },
    });
    if ((res.status === 429 || res.status >= 500) && attempt < 4) {
      const wait = Number(res.headers.get("Retry-After")) || 2 ** attempt;
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    if (!res.ok) throw new Error(`Microsoft Graph ${res.status}: ${(await res.text()).slice(0, 200)}`);
    return res.json() as Promise<T>;
  }
}

type Page<T> = { value: T[]; "@odata.nextLink"?: string };

export type Folder = { id: string; name: string; depth: number; total: number };

type RawFolder = { id: string; displayName: string; childFolderCount: number; totalItemCount: number };

/** All mail folders, flattened in tree order (Inbox, Inbox/Sub, …), up to 4 levels deep. */
export async function listFolders(): Promise<Folder[]> {
  const out: Folder[] = [];
  async function walk(path: string, depth: number) {
    let next: string | undefined = `${path}?$top=100&$select=id,displayName,childFolderCount,totalItemCount`;
    while (next) {
      const page: Page<RawFolder> = await get<Page<RawFolder>>(next);
      for (const f of page.value) {
        out.push({ id: f.id, name: f.displayName, depth, total: f.totalItemCount });
        if (f.childFolderCount > 0 && depth < 3) await walk(`/me/mailFolders/${f.id}/childFolders`, depth + 1);
      }
      next = page["@odata.nextLink"];
    }
  }
  await walk("/me/mailFolders", 0);
  return out;
}

export type Message = {
  id: string;
  subject: string;
  from: string;
  received: string; // ISO
  body: string;
  webLink: string;
};

type RawMessage = {
  id: string;
  subject: string | null;
  receivedDateTime: string;
  webLink: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  body?: { content?: string };
};

/** Messages in a folder received in [since, until), newest first, page by page. */
export async function* listMessages(folderId: string, since: Date, until: Date): AsyncGenerator<Message[]> {
  const filter = `receivedDateTime ge ${since.toISOString()} and receivedDateTime lt ${until.toISOString()}`;
  let next: string | undefined =
    `/me/mailFolders/${folderId}/messages?$filter=${encodeURIComponent(filter)}` +
    `&$orderby=receivedDateTime desc&$top=50&$select=id,subject,from,receivedDateTime,body,webLink`;
  while (next) {
    const page: Page<RawMessage> = await get<Page<RawMessage>>(next);
    yield page.value.map((m) => ({
      id: m.id,
      subject: m.subject ?? "",
      from: m.from?.emailAddress?.address ?? m.from?.emailAddress?.name ?? "",
      received: m.receivedDateTime,
      body: m.body?.content ?? "",
      webLink: m.webLink,
    }));
    next = page["@odata.nextLink"];
  }
}
