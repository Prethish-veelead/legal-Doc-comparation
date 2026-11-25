// Microsoft Graph access with the app registration (client credentials).
import { getConfig } from "./config.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
let token = { value: null, expiresAt: 0 };
let siteId = null;

export class UserError extends Error {
  constructor(message, status = 400) {
    super(message);
    this.status = status;
  }
}

async function getToken() {
  if (token.value && Date.now() < token.expiresAt - 60_000) return token.value;
  const c = getConfig();
  const res = await fetch(`https://login.microsoftonline.com/${c.tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    body: new URLSearchParams({
      client_id: c.clientId,
      client_secret: c.clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  });
  const body = await res.json();
  if (!res.ok) throw new Error(`SharePoint sign-in failed: ${body.error || res.status}`);
  token = { value: body.access_token, expiresAt: Date.now() + body.expires_in * 1000 };
  return token.value;
}

// Calls Graph, retrying on throttling. Returns parsed JSON (or the raw Response when raw=true).
export async function graph(path, { method = "GET", body, raw = false, headers = {} } = {}) {
  const url = path.startsWith("https://") ? path : GRAPH + path;
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${await getToken()}`,
        ...(body ? { "Content-Type": "application/json" } : {}),
        ...headers,
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    if ((res.status === 429 || res.status === 503) && attempt < 4) {
      const wait = Number(res.headers.get("retry-after")) || 2 ** attempt;
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    if (raw) return res;
    if (res.status === 204) return null;
    const json = await res.json().catch(() => ({}));
    if (!res.ok) {
      const err = new Error(`Graph ${method} ${path.split("?")[0]} failed: ${res.status} ${json.error?.message || ""}`);
      err.status = res.status;
      throw err;
    }
    return json;
  }
}

export async function getSiteId() {
  if (siteId) return siteId;
  const c = getConfig();
  siteId = (await graph(`/sites/${c.siteHost}:${c.sitePath}?$select=id`)).id;
  return siteId;
}

// Turns any accepted link form into a server-relative path, or null for sharing links.
function toServerPath(u) {
  const idParam = u.searchParams.get("id"); // .../AllItems.aspx?id=/sites/x/Lib/file.docx
  const p = decodeURIComponent(idParam || u.pathname);
  if (/\/_layouts\//i.test(p) || /^\/:[a-z]:\//i.test(p)) return null; // Doc.aspx / sharing links
  return p;
}

// Downloads a document from the configured site into memory. Never writes to disk.
export async function downloadDocument(documentUrl) {
  const c = getConfig();
  let u;
  try {
    u = new URL(documentUrl);
  } catch {
    throw new UserError("The document link is not a valid URL.");
  }
  if (u.protocol !== "https:" || u.hostname.toLowerCase() !== c.siteHost) {
    throw new UserError(`Only links on ${c.siteHost} are accepted.`);
  }

  const serverPath = toServerPath(u);
  const outside = () => new UserError(`Only documents inside ${c.sitePath} are accepted.`);
  const sid = await getSiteId();
  const drives = (await graph(`/sites/${sid}/drives?$select=id,webUrl`)).value;
  let item;
  if (serverPath) {
    if (!serverPath.toLowerCase().startsWith(c.sitePath.toLowerCase() + "/")) throw outside();
    const segments = serverPath.slice(1).split("/");
    if (segments.some((s) => s === "" || s === "." || s === "..")) throw new UserError("The document link is not valid.");
    const rel = serverPath.slice(c.sitePath.length + 1); // Lib/folder/file.docx
    const drive = drives
      .map((d) => ({ id: d.id, root: decodeURIComponent(new URL(d.webUrl).pathname).slice(c.sitePath.length + 1) }))
      .find((d) => rel.toLowerCase().startsWith(d.root.toLowerCase() + "/"));
    if (!drive) throw new UserError("The link does not point to a document library on this site.");
    const itemPath = rel.slice(drive.root.length + 1);
    item = await graph(
      `/drives/${drive.id}/root:/${itemPath.split("/").map(encodeURIComponent).join("/")}?$select=id,name,size,file,parentReference,webUrl`,
    ).catch((e) => {
      if (e.status === 404) throw new UserError("No document was found at that link.", 404);
      throw e;
    });
  } else {
    // Sharing link or Office Online link: resolve through the shares API.
    const shareId = "u!" + Buffer.from(documentUrl).toString("base64url");
    item = await graph(`/shares/${shareId}/driveItem?$select=id,name,size,file,parentReference,webUrl`).catch((e) => {
      if (e.status === 404 || e.status === 403) throw new UserError("That sharing link could not be opened.", 404);
      throw e;
    });
  }
  // Scope check on the resolved item, whichever way it was found: it must sit in one of this site's libraries.
  if (!drives.some((d) => d.id === item.parentReference?.driveId)) throw outside();

  if (!item.file) throw new UserError("The link points to a folder, not a document.");
  if (item.size > c.maxFileBytes) {
    throw new UserError(`The document is larger than the ${c.maxFileBytes / 1024 / 1024} MB limit.`, 413);
  }

  const res = await graph(`/drives/${item.parentReference.driveId}/items/${item.id}/content`, { raw: true });
  if (!res.ok) throw new Error(`Download failed: ${res.status}`);
  const buffer = Buffer.from(await res.arrayBuffer());
  return {
    buffer,
    fileName: item.name,
    sizeBytes: item.size,
    mimeType: item.file.mimeType,
    driveId: item.parentReference.driveId,
    itemId: item.id,
    webUrl: item.webUrl,
  };
}
